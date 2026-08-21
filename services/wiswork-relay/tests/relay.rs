use axum::response::IntoResponse;
use futures_util::{SinkExt, StreamExt};
use serde_json::{Value, json};
use tokio::net::TcpListener;
use tokio_tungstenite::{
    connect_async,
    tungstenite::{Message, client::IntoClientRequest},
};
use wiswork_relay::{Config, app};

const ORIGIN: &str = "https://office.8-216-134-194.sslip.io";

async fn server() -> String {
    let auth_listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let auth_addr = auth_listener.local_addr().unwrap();
    let auth = axum::Router::new().route(
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
    );
    tokio::spawn(async move { axum::serve(auth_listener, auth).await.unwrap() });
    let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let addr = listener.local_addr().unwrap();
    let config = Config {
        auth_url: format!("http://{auth_addr}/oidc/me"),
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
