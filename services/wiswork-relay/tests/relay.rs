use axum::response::IntoResponse;
use base64::{Engine as _, engine::general_purpose::STANDARD};
use futures_util::{SinkExt, StreamExt};
use jsonwebtoken::{Algorithm, EncodingKey, Header, encode};
use serde::Serialize;
use serde_json::{Value, json};
use tokio::net::TcpListener;
use tokio_tungstenite::{
    connect_async,
    tungstenite::{Message, client::IntoClientRequest},
};
use wiswork_relay::{Config, app};

const ORIGIN: &str = "https://office.8-216-134-194.sslip.io";
const TEST_ISSUER: &str = "https://issuer.example.test";
const TEST_AUDIENCE: &str = "relay-test-client";
const TEST_PRIVATE_KEY_DER: &str = "MIG2AgEAMBAGByqGSM49AgEGBSuBBAAiBIGeMIGbAgEBBDBi2P/vImpNt3oyLprNPTVoaYRvIaJWVxDjsjoRF0YfAmwKUhtjGdj0qh0NWGlbVVOhZANiAAQ1aooGczWALsnMxfjM77d43rpKPqBtQEHPDrizP7fpwi/SbkZ2T/czW8Ye+5Ix3WIYFMM60AKyMtLQXT8V4VjQb0jM9wRkC/JEa0C50q9dk8APUPfbJMDbcsqcyW0bl2w=";

async fn server() -> String {
    let auth_listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let auth_addr = auth_listener.local_addr().unwrap();
    let auth = axum::Router::new()
        .route(
            "/oidc/me",
            axum::routing::get(|headers: axum::http::HeaderMap| async move {
            if headers.get("authorization").and_then(|v| v.to_str().ok())
                == Some("Bearer valid-test-token")
            {
                (
                    axum::http::StatusCode::OK,
                    axum::Json(json!({"sub":"test-user"})),
                )
                    .into_response()
            } else {
                axum::http::StatusCode::UNAUTHORIZED.into_response()
            }
            }),
        )
        .route(
            "/oidc/jwks",
            axum::routing::get(|| async {
                axum::Json(json!({"keys":[{"kty":"EC","crv":"P-384","kid":"test-key","use":"sig","alg":"ES384","x":"NWqKBnM1gC7JzMX4zO-3eN66Sj6gbUBBzw64sz-36cIv0m5Gdk_3M1vGHvuSMd1i","y":"GBTDOtACsjLS0F0_FeFY0G9IzPcEZAvyRGtAudKvXZPAD1D32yTA23LKnMltG5ds"}]}))
            }),
        );
    tokio::spawn(async move { axum::serve(auth_listener, auth).await.unwrap() });
    let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let addr = listener.local_addr().unwrap();
    let config = Config {
        auth_url: format!("http://{auth_addr}/oidc/me"),
        jwks_url: format!("http://{auth_addr}/oidc/jwks"),
        issuer: TEST_ISSUER.into(),
        audience: TEST_AUDIENCE.into(),
        ..Config::default()
    };
    tokio::spawn(async move {
        axum::serve(
            listener,
            app(config).into_make_service_with_connect_info::<std::net::SocketAddr>(),
        )
        .await
        .unwrap()
    });
    format!("ws://{addr}/office-relay")
}

async fn socket(
    url: &str,
    origin: &str,
) -> tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>> {
    let mut request = url.into_client_request().unwrap();
    request
        .headers_mut()
        .insert("origin", origin.parse().unwrap());
    connect_async(request).await.unwrap().0
}

async fn pc_socket(
    url: &str,
) -> tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>> {
    let mut request = url.into_client_request().unwrap();
    request
        .headers_mut()
        .insert("authorization", "Bearer valid-test-token".parse().unwrap());
    connect_async(request).await.unwrap().0
}

async fn pc_socket_with_token(
    url: &str,
    token: &str,
) -> Result<
    tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>>,
    tokio_tungstenite::tungstenite::Error,
> {
    let mut request = url.into_client_request().unwrap();
    request
        .headers_mut()
        .insert("authorization", format!("Bearer {token}").parse().unwrap());
    connect_async(request).await.map(|value| value.0)
}

#[derive(Serialize)]
struct TestClaims<'a> {
    sub: &'a str,
    iss: &'a str,
    aud: &'a str,
    exp: u64,
}

fn id_token(issuer: &str, audience: &str, expires_at: u64) -> String {
    let mut header = Header::new(Algorithm::ES384);
    header.kid = Some("test-key".into());
    encode(
        &header,
        &TestClaims {
            sub: "jwt-test-user",
            iss: issuer,
            aud: audience,
            exp: expires_at,
        },
        &EncodingKey::from_ec_der(&STANDARD.decode(TEST_PRIVATE_KEY_DER).unwrap()),
    )
    .unwrap()
}

async fn send(
    ws: &mut (impl SinkExt<Message, Error = tokio_tungstenite::tungstenite::Error> + Unpin),
    value: Value,
) {
    ws.send(Message::Text(value.to_string().into()))
        .await
        .unwrap();
}

async fn recv(
    ws: &mut (impl StreamExt<Item = Result<Message, tokio_tungstenite::tungstenite::Error>> + Unpin),
) -> Value {
    let Message::Text(text) = ws.next().await.unwrap().unwrap() else {
        panic!("expected text")
    };
    serde_json::from_str(&text).unwrap()
}

#[tokio::test]
async fn health_and_exact_origin() {
    let url = server().await;
    let health = reqwest::Client::builder()
        .no_proxy()
        .build()
        .unwrap()
        .get(url.replace("ws://", "http://") + "/health")
        .send()
        .await
        .unwrap();
    assert_eq!(health.status(), 200);
    assert!(connect_async(&url).await.is_err());
    let mut invalid_pc = url.as_str().into_client_request().unwrap();
    invalid_pc
        .headers_mut()
        .insert("authorization", "Bearer invalid".parse().unwrap());
    assert!(connect_async(invalid_pc).await.is_err());
    let mut no_origin = pc_socket(&url).await;
    send(
        &mut no_origin,
        json!({"version":1,"type":"office.create","host":"Word"}),
    )
    .await;
    assert_eq!(recv(&mut no_origin).await["code"], "role_not_allowed");
    assert!(
        connect_async({
            let mut r = url.as_str().into_client_request().unwrap();
            r.headers_mut()
                .insert("origin", "https://evil.example".parse().unwrap());
            r
        })
        .await
        .is_err()
    );
}

#[tokio::test]
async fn authenticates_gateway_id_tokens_with_fixed_oidc_jwks() {
    let url = server().await;
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_secs();
    let valid = id_token(TEST_ISSUER, TEST_AUDIENCE, now + 300);
    let mut pc = pc_socket_with_token(&url, &valid).await.unwrap();
    send(
        &mut pc,
        json!({"version":1,"type":"pc.claim","verification_code":"999999"}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["code"], "invalid_code");

    let wrong_audience = id_token(TEST_ISSUER, "other-client", now + 300);
    assert!(pc_socket_with_token(&url, &wrong_audience).await.is_err());

    let wrong_issuer = id_token("https://attacker.example.test", TEST_AUDIENCE, now + 300);
    assert!(pc_socket_with_token(&url, &wrong_issuer).await.is_err());

    let expired = id_token(TEST_ISSUER, TEST_AUDIENCE, now - 120);
    assert!(pc_socket_with_token(&url, &expired).await.is_err());
}

#[tokio::test]
async fn rate_limits_claims_across_reconnects_and_pairings_per_connection() {
    let url = server().await;
    for attempt in 0..6 {
        let mut pc = pc_socket(&url).await;
        send(
            &mut pc,
            json!({"version":1,"type":"pc.claim","verification_code":"999999"}),
        )
        .await;
        let response = recv(&mut pc).await;
        assert_eq!(
            response["code"],
            if attempt < 5 {
                "invalid_code"
            } else {
                "claim_limit"
            }
        );
    }

    let mut office = socket(&url, ORIGIN).await;
    for _ in 0..4 {
        send(
            &mut office,
            json!({"version":1,"type":"office.create","host":"Word"}),
        )
        .await;
        assert_eq!(recv(&mut office).await["type"], "office.created");
    }
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"Word"}),
    )
    .await;
    assert_eq!(recv(&mut office).await["code"], "pairing_limit");
}

#[tokio::test]
async fn pc_disconnect_releases_claim_but_preserves_code_and_attempts() {
    let url = server().await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"Excel"}),
    )
    .await;
    let created = recv(&mut office).await;
    let code = created["verification_code"].as_str().unwrap();

    let mut first = pc_socket(&url).await;
    send(
        &mut first,
        json!({"version":1,"type":"pc.claim","verification_code":code}),
    )
    .await;
    assert_eq!(recv(&mut first).await["type"], "pc.claimed");
    first.close(None).await.unwrap();
    assert_eq!(recv(&mut office).await["type"], "office.pc_offline");

    let mut second = pc_socket(&url).await;
    send(
        &mut second,
        json!({"version":1,"type":"pc.claim","verification_code":code}),
    )
    .await;
    assert_eq!(recv(&mut second).await["type"], "pc.claimed");
}

#[tokio::test]
async fn pairs_only_after_claim_and_approval_then_forwards_and_cancels() {
    let url = server().await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"Word"}),
    )
    .await;
    let created = recv(&mut office).await;
    assert_eq!(created["type"], "office.created");
    let code = created["verification_code"].as_str().unwrap();
    assert_eq!(code.len(), 6);

    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.claim","verification_code":code}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    assert_eq!(claimed["host"], "Word");
    send(
        &mut pc,
        json!({"version":1,"type":"pc.approve","pairing_id":claimed["pairing_id"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert_ne!(pc_ready["capability"], office_ready["capability"]);

    let cap = office_ready["capability"].clone();
    let sid = office_ready["session_id"].clone();
    send(&mut office, json!({"version":1,"type":"office.request","session_id":sid,"capability":cap,"request_id":"r1","body":{"messages":[]}})).await;
    assert_eq!(recv(&mut pc).await["type"], "relay.request");
    send(&mut office, json!({"version":1,"type":"office.cancel","session_id":sid,"capability":cap,"request_id":"r1"})).await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");

    send(&mut office, json!({"version":1,"type":"office.request","session_id":sid,"capability":cap,"request_id":"r2","body":{"messages":[]}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "r2");
    let pc_cap = pc_ready["capability"].clone();
    send(&mut pc, json!({"version":1,"type":"pc.start","session_id":sid,"capability":pc_cap,"request_id":"r2","status":200,"content_type":"text/event-stream"})).await;
    assert_eq!(recv(&mut office).await["type"], "relay.start");
    send(&mut pc, json!({"version":1,"type":"pc.chunk","session_id":sid,"capability":pc_cap,"request_id":"r2","sequence":0,"data":"b2s="})).await;
    assert_eq!(recv(&mut office).await["data"], "b2s=");
    send(&mut pc, json!({"version":1,"type":"pc.done","session_id":sid,"capability":pc_cap,"request_id":"r2"})).await;
    assert_eq!(recv(&mut office).await["type"], "relay.done");

    let boundary_body = "x".repeat(256 * 1024 - 2);
    send(&mut office, json!({"version":1,"type":"office.request","session_id":sid,"capability":cap,"request_id":"boundary","body":boundary_body})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "boundary");
    send(&mut office, json!({"version":1,"type":"office.cancel","session_id":sid,"capability":cap,"request_id":"boundary"})).await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");

    let oversized_body = "x".repeat(256 * 1024 - 1);
    send(&mut office, json!({"version":1,"type":"office.request","session_id":sid,"capability":cap,"request_id":"oversized","body":oversized_body})).await;
    assert_eq!(recv(&mut office).await["code"], "request_too_large");
}

#[tokio::test]
async fn v2_negotiates_exact_capabilities_and_denies_unnegotiated_requests() {
    let url = server().await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.create","host":"Word","capabilities":["agent.v1","web-search.v1","web-fetch.v1"]}),
    )
    .await;
    let created = recv(&mut office).await;
    assert_eq!(created["version"], 2);
    let code = created["verification_code"].as_str().unwrap();

    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.claim","verification_code":code,"capabilities":["agent.v1","web-search.v1","image-search.v1"]}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    assert_eq!(
        claimed["capabilities"],
        json!(["agent.v1", "web-search.v1"])
    );
    send(
        &mut pc,
        json!({"version":2,"type":"pc.approve","pairing_id":claimed["pairing_id"],"capabilities":["agent.v1","web-search.v1"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert_eq!(
        office_ready["capabilities"],
        json!(["agent.v1", "web-search.v1"])
    );

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"web_request_1","capability_name":"web-fetch.v1","body":{"url":"https://example.com"}}),
    )
    .await;
    assert_eq!(recv(&mut office).await["code"], "capability_not_negotiated");

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"web_request_2","capability_name":"web-search.v1","body":{"query":"office agents","max_results":5}}),
    )
    .await;
    let forwarded = recv(&mut pc).await;
    assert_eq!(forwarded["version"], 2);
    assert_eq!(forwarded["capability_name"], "web-search.v1");
    assert_eq!(
        forwarded["body"],
        json!({"query":"office agents","max_results":5})
    );
    assert_eq!(pc_ready["capabilities"], office_ready["capabilities"]);
}

#[tokio::test]
async fn rejects_extra_keys_malformed_binary_and_pc_spoofed_origin() {
    let url = server().await;
    let mut spoof = socket(&url, ORIGIN).await;
    send(
        &mut spoof,
        json!({"version":1,"type":"pc.claim","verification_code":"123456"}),
    )
    .await;
    assert_eq!(recv(&mut spoof).await["code"], "role_not_allowed");
    let mut invalid = socket(&url, ORIGIN).await;
    send(
        &mut invalid,
        json!({"version":2,"type":"office.create","host":"Word","extra":1}),
    )
    .await;
    assert_eq!(recv(&mut invalid).await["code"], "invalid_frame");
    let mut binary = socket(&url, ORIGIN).await;
    binary
        .send(Message::Binary(vec![1, 2, 3].into()))
        .await
        .unwrap();
    assert_eq!(recv(&mut binary).await["code"], "binary_not_supported");
}
