use axum::response::IntoResponse;
use base64::{
    Engine as _,
    engine::general_purpose::{STANDARD, URL_SAFE_NO_PAD},
};
use futures_util::{SinkExt, StreamExt};
use jsonwebtoken::{Algorithm, EncodingKey, Header, encode};
use p256::ecdsa::{Signature, SigningKey, signature::Signer};
use serde::Serialize;
use serde_json::{Value, json};
use std::{
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Duration,
};
use tokio::net::TcpListener;
use tokio_tungstenite::{
    connect_async,
    tungstenite::{Message, client::IntoClientRequest},
};
use wiswork_relay::{Config, app, try_app};

type TestSocket =
    tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>>;

static NEXT_DATABASE: AtomicU64 = AtomicU64::new(1);

fn database_path(label: &str) -> PathBuf {
    std::env::temp_dir().join(format!(
        "wiswork-relay-{}-{}-{label}.sqlite",
        std::process::id(),
        NEXT_DATABASE.fetch_add(1, Ordering::Relaxed)
    ))
}

#[test]
fn default_request_budget_supports_complex_agent_turns() {
    assert_eq!(Config::default().request_ttl, Duration::from_secs(300));
    assert_eq!(Config::default().session_ttl, Duration::from_secs(1800));
    assert_eq!(
        Config::default().auth_url,
        "https://auth.wispaper.ai/oidc/me"
    );
    assert_eq!(
        Config::default().jwks_url,
        "https://auth.wispaper.ai/oidc/jwks"
    );
    assert_eq!(Config::default().issuer, "https://auth.wispaper.ai/oidc");
    assert_eq!(Config::default().audience, "i9au2rbqzktme4runr9gy");
}

const ORIGIN: &str = "https://office.8-216-134-194.sslip.io";
const TEST_ISSUER: &str = "https://issuer.example.test";
const TEST_AUDIENCE: &str = "relay-test-client";
const TEST_PRIVATE_KEY_DER: &str = "MIG2AgEAMBAGByqGSM49AgEGBSuBBAAiBIGeMIGbAgEBBDBi2P/vImpNt3oyLprNPTVoaYRvIaJWVxDjsjoRF0YfAmwKUhtjGdj0qh0NWGlbVVOhZANiAAQ1aooGczWALsnMxfjM77d43rpKPqBtQEHPDrizP7fpwi/SbkZ2T/czW8Ye+5Ix3WIYFMM60AKyMtLQXT8V4VjQb0jM9wRkC/JEa0C50q9dk8APUPfbJMDbcsqcyW0bl2w=";

#[test]
fn deployment_bounds_diagnostic_journal_retention() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    let unit = std::fs::read_to_string(root.join("deploy/wiswork-relay.service")).unwrap();
    let journal = std::fs::read_to_string(root.join("deploy/journald@wiswork-relay.conf")).unwrap();
    assert!(unit.contains("LogNamespace=wiswork-relay"));
    assert!(unit.contains("StateDirectory=wiswork-relay"));
    assert!(unit.contains("StateDirectoryMode=0700"));
    assert!(unit.contains("UMask=0077"));
    assert!(unit.contains("WISWORK_RELAY_BINDING_DB=/var/lib/wiswork-relay/bindings.sqlite"));
    assert!(journal.contains("MaxRetentionSec=7day"));
    assert!(journal.contains("SystemMaxUse=64M"));
}

async fn server_with_all_limits(
    session_ttls: Option<(Duration, Duration)>,
    max_global_claims: Option<u32>,
    diagnostic_rate: Option<(u8, Duration)>,
) -> String {
    let auth_listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let auth_addr = auth_listener.local_addr().unwrap();
    let auth = axum::Router::new()
        .route(
            "/oidc/me",
            axum::routing::get(|headers: axum::http::HeaderMap| async move {
            let subject = match headers.get("authorization").and_then(|v| v.to_str().ok()) {
                Some("Bearer valid-test-token") => Some("test-user"),
                Some("Bearer legitimate-test-token") => Some("legitimate-user"),
                _ => None,
            };
            if let Some(subject) = subject {
                (
                    axum::http::StatusCode::OK,
                    axum::Json(json!({"sub":subject})),
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
    let mut config = Config {
        auth_url: format!("http://{auth_addr}/oidc/me"),
        jwks_url: format!("http://{auth_addr}/oidc/jwks"),
        issuer: TEST_ISSUER.into(),
        audience: TEST_AUDIENCE.into(),
        ..Config::default()
    };
    if let Some((idle, maximum)) = session_ttls {
        config.session_ttl = idle;
        config.session_max_ttl = maximum;
    }
    if let Some(maximum) = max_global_claims {
        config.max_global_claims = maximum;
    }
    if let Some((maximum, window)) = diagnostic_rate {
        config.max_diagnostics_per_window = maximum;
        config.diagnostic_window = window;
    }
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

async fn server_with_binding_database(path: &Path, challenge_ttl: Option<Duration>) -> String {
    let auth_listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let auth_addr = auth_listener.local_addr().unwrap();
    let auth = axum::Router::new().route(
        "/oidc/me",
        axum::routing::get(|headers: axum::http::HeaderMap| async move {
            let subject = match headers.get("authorization").and_then(|v| v.to_str().ok()) {
                Some("Bearer valid-test-token") => Some("test-user"),
                Some("Bearer legitimate-test-token") => Some("legitimate-user"),
                _ => None,
            };
            if let Some(subject) = subject {
                (
                    axum::http::StatusCode::OK,
                    axum::Json(json!({"sub":subject})),
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
    let mut config = Config {
        auth_url: format!("http://{auth_addr}/oidc/me"),
        binding_database: Some(path.to_owned()),
        ..Config::default()
    };
    if let Some(challenge_ttl) = challenge_ttl {
        config.resume_challenge_ttl = challenge_ttl;
    }
    let router = try_app(config).unwrap();
    tokio::spawn(async move {
        axum::serve(
            listener,
            router.into_make_service_with_connect_info::<std::net::SocketAddr>(),
        )
        .await
        .unwrap()
    });
    format!("ws://{addr}/office-relay")
}

async fn server_with_limits(
    session_ttls: Option<(Duration, Duration)>,
    max_global_claims: Option<u32>,
) -> String {
    server_with_all_limits(session_ttls, max_global_claims, None).await
}

async fn server_with_diagnostic_rate(maximum: u8, window: Duration) -> String {
    server_with_all_limits(None, None, Some((maximum, window))).await
}

async fn server_with_session_ttls(session_ttls: Option<(Duration, Duration)>) -> String {
    server_with_limits(session_ttls, None).await
}

async fn server() -> String {
    server_with_session_ttls(None).await
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
    Box<tokio_tungstenite::tungstenite::Error>,
> {
    let mut request = url.into_client_request().unwrap();
    request
        .headers_mut()
        .insert("authorization", format!("Bearer {token}").parse().unwrap());
    connect_async(request)
        .await
        .map(|value| value.0)
        .map_err(Box::new)
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

async fn approved_v2_session(
    url: &str,
) -> (
    tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>>,
    tokio_tungstenite::WebSocketStream<tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>>,
    Value,
    Value,
) {
    let mut office = socket(url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.create","host":"Word","capabilities":["agent.v1"]}),
    )
    .await;
    let created = recv(&mut office).await;
    let mut pc = pc_socket(url).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.claim","verification_code":created["verification_code"],"capabilities":["agent.v1"]}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.approve","pairing_id":claimed["pairing_id"],"capabilities":["agent.v1"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    (office, pc, office_ready, pc_ready)
}

async fn begin_enrollment(url: &str, signing_key: &SigningKey) -> (TestSocket, TestSocket, Value) {
    let public_key = URL_SAFE_NO_PAD.encode(signing_key.verifying_key().to_encoded_point(false));
    let mut office = socket(url, ORIGIN).await;
    send(
        &mut office,
        json!({
            "version": 2,
            "type": "office.create",
            "host": "Word",
            "capabilities": ["agent.v1"],
            "features": ["pairing-resume.v1"],
            "binding_public_key": public_key,
        }),
    )
    .await;

    let created = recv(&mut office).await;
    assert_eq!(created["type"], "office.created");
    assert_eq!(created["features"], json!(["pairing-resume.v1"]));
    assert!(created.get("binding_id").is_none());
    let mut pc = pc_socket(url).await;
    send(
        &mut pc,
        json!({
            "version": 2,
            "type": "pc.claim",
            "verification_code": created["verification_code"],
            "capabilities": ["agent.v1"],
            "features": ["pairing-resume.v1"],
        }),
    )
    .await;
    let claimed = recv(&mut pc).await;
    assert_eq!(claimed["type"], "pc.claimed");
    assert_eq!(claimed["features"], json!(["pairing-resume.v1"]));
    (office, pc, claimed)
}

async fn enroll(
    url: &str,
    signing_key: &SigningKey,
) -> (TestSocket, TestSocket, Value, Value, String) {
    let (mut office, mut pc, claimed) = begin_enrollment(url, signing_key).await;
    send(
        &mut pc,
        json!({
            "version": 2,
            "type": "pc.approve",
            "pairing_id": claimed["pairing_id"],
            "capabilities": ["agent.v1"],
            "features": ["pairing-resume.v1"],
        }),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert_eq!(pc_ready["features"], json!(["pairing-resume.v1"]));
    assert_eq!(office_ready["features"], json!(["pairing-resume.v1"]));
    assert_eq!(pc_ready["binding_id"], office_ready["binding_id"]);
    let binding_id = office_ready["binding_id"].as_str().unwrap().to_owned();
    (office, pc, office_ready, pc_ready, binding_id)
}

async fn start_office_resume(
    url: &str,
    signing_key: &SigningKey,
    binding_id: &str,
) -> (TestSocket, Value) {
    let mut office = socket(url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]}),
    )
    .await;
    let challenge = recv(&mut office).await;
    assert_eq!(challenge["type"], "office.challenge");
    assert!(challenge["expires_in"].as_u64().unwrap() <= 30);
    let challenge_value = challenge["challenge"].as_str().unwrap();
    let transcript =
        format!("wiswork-office-resume-v1\n{binding_id}\n{challenge_value}\n{ORIGIN}\nWord");
    let signature: Signature = signing_key.sign(transcript.as_bytes());
    send(
        &mut office,
        json!({"version":2,"type":"office.prove","binding_id":binding_id,"challenge":challenge_value,"signature":URL_SAFE_NO_PAD.encode(signature.to_bytes())}),
    )
    .await;
    (office, challenge)
}

async fn resume(
    url: &str,
    signing_key: &SigningKey,
    binding_id: &str,
) -> (TestSocket, TestSocket, Value, Value) {
    let (mut office, _) = start_office_resume(url, signing_key, binding_id).await;
    assert_eq!(recv(&mut office).await["type"], "office.waiting_for_pc");
    let mut pc = pc_socket(url).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["agent.v1"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert_eq!(pc_ready["type"], "pc.approved");
    assert_eq!(office_ready["type"], "office.approved");
    assert!(pc_ready.get("binding_id").is_none());
    assert!(office_ready.get("features").is_none());
    (office, pc, office_ready, pc_ready)
}

#[tokio::test]
async fn enhanced_enrollment_is_opt_in_and_requires_explicit_approval() {
    let path = database_path("approval");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[7_u8; 32]).unwrap();
    let (mut office, mut pc, claimed) = begin_enrollment(&url, &signing_key).await;
    let connection = rusqlite::Connection::open(&path).unwrap();
    let count: i64 = connection
        .query_row("SELECT count(*) FROM durable_bindings", [], |row| {
            row.get(0)
        })
        .unwrap();
    assert_eq!(count, 0);
    send(
        &mut pc,
        json!({"version":2,"type":"pc.approve","pairing_id":claimed["pairing_id"],"capabilities":["agent.v1"],"features":["pairing-resume.v1"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert_eq!(pc_ready["binding_id"], office_ready["binding_id"]);
    let count: i64 = connection
        .query_row("SELECT count(*) FROM durable_bindings", [], |row| {
            row.get(0)
        })
        .unwrap();
    assert_eq!(count, 1);
}

#[tokio::test]
async fn resumes_after_relay_restart_with_fresh_ephemeral_session() {
    let path = database_path("restart");
    let first_url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[8_u8; 32]).unwrap();
    let (mut first_office, mut first_pc, initial_office, initial_pc, binding_id) =
        enroll(&first_url, &signing_key).await;
    first_office.close(None).await.unwrap();
    first_pc.close(None).await.unwrap();

    let restarted_url = server_with_binding_database(&path, None).await;
    let (_office, _pc, resumed_office, resumed_pc) =
        resume(&restarted_url, &signing_key, &binding_id).await;
    assert_ne!(resumed_office["session_id"], initial_office["session_id"]);
    assert_ne!(resumed_office["capability"], initial_office["capability"]);
    assert_ne!(resumed_pc["capability"], initial_pc["capability"]);
}

#[tokio::test]
async fn concurrent_documents_resume_to_distinct_sessions() {
    let path = database_path("concurrent");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[9_u8; 32]).unwrap();
    let (_office, _pc, _, _, binding_id) = enroll(&url, &signing_key).await;
    let (mut office_one, _) = start_office_resume(&url, &signing_key, &binding_id).await;
    let (mut office_two, _) = start_office_resume(&url, &signing_key, &binding_id).await;
    assert_eq!(recv(&mut office_one).await["type"], "office.waiting_for_pc");
    assert_eq!(recv(&mut office_two).await["type"], "office.waiting_for_pc");

    let mut pc_one = pc_socket(&url).await;
    send(
        &mut pc_one,
        json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["agent.v1"]}),
    )
    .await;
    let first_pc_ready = recv(&mut pc_one).await;
    let first_office_ready = tokio::select! {
        value = recv(&mut office_one) => value,
        value = recv(&mut office_two) => value,
    };
    let mut pc_two = pc_socket(&url).await;
    send(
        &mut pc_two,
        json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["agent.v1"]}),
    )
    .await;
    let second_pc_ready = recv(&mut pc_two).await;
    let second_office_ready = tokio::select! {
        value = recv(&mut office_one) => value,
        value = recv(&mut office_two) => value,
    };
    assert_ne!(first_pc_ready["session_id"], second_pc_ready["session_id"]);
    assert_ne!(
        first_office_ready["capability"],
        second_office_ready["capability"]
    );
}

#[tokio::test]
async fn pc_can_wait_before_office_proves_the_binding() {
    let path = database_path("pc-first");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[15_u8; 32]).unwrap();
    let (_office, _pc, _, _, binding_id) = enroll(&url, &signing_key).await;
    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["agent.v1"]}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["type"], "pc.waiting_for_office");
    let (mut office, _) = start_office_resume(&url, &signing_key, &binding_id).await;
    let office_ready = recv(&mut office).await;
    let pc_ready = recv(&mut pc).await;
    assert_eq!(office_ready["type"], "office.approved");
    assert_eq!(pc_ready["type"], "pc.approved");
    assert_eq!(office_ready["session_id"], pc_ready["session_id"]);
}

#[tokio::test]
async fn rejects_forged_replayed_and_expired_connection_bound_proofs() {
    let path = database_path("proofs");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[10_u8; 32]).unwrap();
    let (_office, _pc, _, _, binding_id) = enroll(&url, &signing_key).await;

    let mut challenge_owner = socket(&url, ORIGIN).await;
    send(&mut challenge_owner, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
    let stolen_challenge = recv(&mut challenge_owner).await;
    let challenge_value = stolen_challenge["challenge"].as_str().unwrap();
    let transcript =
        format!("wiswork-office-resume-v1\n{binding_id}\n{challenge_value}\n{ORIGIN}\nWord");
    let signature: Signature = signing_key.sign(transcript.as_bytes());
    let mut different_connection = socket(&url, ORIGIN).await;
    send(&mut different_connection, json!({"version":2,"type":"office.prove","binding_id":binding_id,"challenge":challenge_value,"signature":URL_SAFE_NO_PAD.encode(signature.to_bytes())})).await;
    assert_eq!(
        recv(&mut different_connection).await["code"],
        "invalid_proof"
    );

    let forged_key = SigningKey::from_slice(&[11_u8; 32]).unwrap();
    let (mut forged_office, _challenge) = start_office_resume(&url, &forged_key, &binding_id).await;
    assert_eq!(recv(&mut forged_office).await["code"], "invalid_proof");

    let (mut proved_office, challenge) = start_office_resume(&url, &signing_key, &binding_id).await;
    assert_eq!(
        recv(&mut proved_office).await["type"],
        "office.waiting_for_pc"
    );
    let challenge_value = challenge["challenge"].as_str().unwrap();
    let transcript =
        format!("wiswork-office-resume-v1\n{binding_id}\n{challenge_value}\n{ORIGIN}\nWord");
    let signature: Signature = signing_key.sign(transcript.as_bytes());
    send(&mut proved_office, json!({"version":2,"type":"office.prove","binding_id":binding_id,"challenge":challenge_value,"signature":URL_SAFE_NO_PAD.encode(signature.to_bytes())})).await;
    assert_eq!(recv(&mut proved_office).await["code"], "invalid_proof");

    let expiring_url = server_with_binding_database(&path, Some(Duration::from_millis(5))).await;
    let mut expired_office = socket(&expiring_url, ORIGIN).await;
    send(&mut expired_office, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
    let expired_challenge = recv(&mut expired_office).await;
    tokio::time::sleep(Duration::from_millis(15)).await;
    let challenge_value = expired_challenge["challenge"].as_str().unwrap();
    let transcript =
        format!("wiswork-office-resume-v1\n{binding_id}\n{challenge_value}\n{ORIGIN}\nWord");
    let signature: Signature = signing_key.sign(transcript.as_bytes());
    send(&mut expired_office, json!({"version":2,"type":"office.prove","binding_id":binding_id,"challenge":challenge_value,"signature":URL_SAFE_NO_PAD.encode(signature.to_bytes())})).await;
    assert_eq!(recv(&mut expired_office).await["code"], "challenge_expired");
}

#[tokio::test]
async fn bounds_resume_challenges_per_connection_and_ip() {
    let path = database_path("resume-limits");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[14_u8; 32]).unwrap();
    let (_office, _pc, _, _, binding_id) = enroll(&url, &signing_key).await;

    let mut same_connection = socket(&url, ORIGIN).await;
    for expected in ["office.challenge", "relay.error"] {
        send(&mut same_connection, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
        let response = recv(&mut same_connection).await;
        assert_eq!(response["type"], expected);
        if expected == "relay.error" {
            assert_eq!(response["code"], "resume_limit");
        }
    }

    for _ in 2..20 {
        let mut office = socket(&url, ORIGIN).await;
        send(&mut office, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
        assert_eq!(recv(&mut office).await["type"], "office.challenge");
        office.close(None).await.unwrap();
    }
    let mut limited = socket(&url, ORIGIN).await;
    send(&mut limited, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
    assert_eq!(recv(&mut limited).await["code"], "resume_rate_limited");
}

#[tokio::test]
async fn resume_fails_closed_for_wrong_subject_host_and_capabilities() {
    let path = database_path("scope");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[12_u8; 32]).unwrap();
    let (_office, _pc, _, _, binding_id) = enroll(&url, &signing_key).await;

    let mut wrong_host = socket(&url, ORIGIN).await;
    send(&mut wrong_host, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Excel","capabilities":["agent.v1"]})).await;
    assert_eq!(recv(&mut wrong_host).await["code"], "binding_unavailable");

    let mut wrong_subject = pc_socket_with_token(&url, "legitimate-test-token")
        .await
        .unwrap();
    send(
        &mut wrong_subject,
        json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["agent.v1"]}),
    )
    .await;
    assert_eq!(
        recv(&mut wrong_subject).await["code"],
        "binding_unavailable"
    );

    let mut wrong_capability = pc_socket(&url).await;
    send(&mut wrong_capability, json!({"version":2,"type":"pc.resume","binding_id":binding_id,"capabilities":["web-fetch.v1"]})).await;
    assert_eq!(
        recv(&mut wrong_capability).await["code"],
        "capability_not_negotiated"
    );
}

#[tokio::test]
async fn revocation_terminates_sessions_and_prevents_future_resume() {
    let path = database_path("revoke");
    let url = server_with_binding_database(&path, None).await;
    let signing_key = SigningKey::from_slice(&[13_u8; 32]).unwrap();
    let (mut office, mut pc, _, _, binding_id) = enroll(&url, &signing_key).await;
    let mut revoker = pc_socket(&url).await;
    send(
        &mut revoker,
        json!({"version":2,"type":"pc.revoke_binding","binding_id":binding_id}),
    )
    .await;
    assert_eq!(recv(&mut revoker).await["type"], "pc.binding_revoked");
    assert_eq!(recv(&mut office).await["code"], "session_revoked");
    assert_eq!(recv(&mut pc).await["code"], "session_revoked");

    let mut resume_office = socket(&url, ORIGIN).await;
    send(&mut resume_office, json!({"version":2,"type":"office.resume","binding_id":binding_id,"host":"Word","capabilities":["agent.v1"]})).await;
    assert_eq!(
        recv(&mut resume_office).await["code"],
        "binding_unavailable"
    );
}

#[test]
fn unknown_database_schema_fails_without_mutation() {
    let path = database_path("future-schema");
    let connection = rusqlite::Connection::open(&path).unwrap();
    connection
        .execute_batch("PRAGMA user_version = 2;")
        .unwrap();
    drop(connection);
    let result = try_app(Config {
        binding_database: Some(path.clone()),
        ..Config::default()
    });
    assert!(result.is_err());
    let connection = rusqlite::Connection::open(path).unwrap();
    let version: i64 = connection
        .query_row("PRAGMA user_version", [], |row| row.get(0))
        .unwrap();
    assert_eq!(version, 2);
}

#[test]
fn malformed_current_database_schema_fails_startup() {
    let path = database_path("malformed-schema");
    let connection = rusqlite::Connection::open(&path).unwrap();
    connection
        .execute_batch("PRAGMA user_version = 1;")
        .unwrap();
    drop(connection);
    assert!(
        try_app(Config {
            binding_database: Some(path),
            ..Config::default()
        })
        .is_err()
    );
}

#[cfg(unix)]
#[tokio::test]
async fn binding_database_file_is_owner_only() {
    use std::os::unix::fs::PermissionsExt;
    let path = database_path("permissions");
    let _router = try_app(Config {
        binding_database: Some(path.clone()),
        ..Config::default()
    })
    .unwrap();
    assert_eq!(
        std::fs::metadata(path).unwrap().permissions().mode() & 0o777,
        0o600
    );
}

#[tokio::test]
async fn persistent_relay_keeps_legacy_v2_frames_exact_and_features_non_callable() {
    let path = database_path("legacy");
    let url = server_with_binding_database(&path, None).await;
    let (mut legacy_office, mut legacy_pc, office_ready, pc_ready) =
        approved_v2_session(&url).await;
    let mut office_keys: Vec<_> = office_ready
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    office_keys.sort_unstable();
    assert_eq!(
        office_keys,
        [
            "capabilities",
            "capability",
            "expires_in",
            "session_id",
            "type",
            "version"
        ]
    );
    assert!(pc_ready.get("features").is_none());
    send(&mut legacy_office, json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"control-is-not-data","capability_name":"pairing-resume.v1","body":{}})).await;
    assert_eq!(
        recv(&mut legacy_office).await["code"],
        "capability_not_negotiated"
    );
    assert!(
        tokio::time::timeout(Duration::from_millis(20), recv(&mut legacy_pc))
            .await
            .is_err()
    );
}

#[tokio::test]
async fn new_pc_uses_empty_feature_intersection_with_legacy_office() {
    let path = database_path("empty-feature-intersection");
    let url = server_with_binding_database(&path, None).await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.create","host":"Word","capabilities":["agent.v1"]}),
    )
    .await;
    let created = recv(&mut office).await;
    let mut pc = pc_socket(&url).await;
    send(&mut pc, json!({"version":2,"type":"pc.negotiate","verification_code":created["verification_code"],"capabilities":["agent.v1"],"features":["pairing-resume.v1"]})).await;
    let negotiated = recv(&mut pc).await;
    assert_eq!(negotiated["features"], json!([]));
    send(&mut pc, json!({"version":2,"type":"pc.claim","verification_code":created["verification_code"],"capabilities":["agent.v1"],"features":[]})).await;
    let claimed = recv(&mut pc).await;
    assert_eq!(claimed["features"], json!([]));
    send(&mut pc, json!({"version":2,"type":"pc.approve","pairing_id":claimed["pairing_id"],"capabilities":["agent.v1"],"features":[]})).await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    assert!(pc_ready.get("features").is_none());
    assert!(pc_ready.get("binding_id").is_none());
    assert!(office_ready.get("features").is_none());
    assert!(office_ready.get("binding_id").is_none());
}

fn diagnostic(ready: &Value, event_id: &str) -> Value {
    json!({
        "version": 2,
        "type": "office.diagnostic",
        "session_id": ready["session_id"],
        "capability": ready["capability"],
        "event_id": event_id,
        "trace_id": "b74ed23d-13a7-4a43-9f86-8d50c7279fa5",
        "timestamp_ms": 1_787_373_000_000_u64,
        "host": "word",
        "platform": "mac",
        "build": "taskpane-BrwOZlpi",
        "tool": "word_write_document",
        "phase": "write",
        "outcome": "failed",
        "error_code": "office_write_failed",
        "office_error_code": "InvalidArgument",
        "office_error_name": "OfficeExtension.Error",
        "office_error_location": "Body.insertText",
        "duration_ms": 42,
        "requirement_sets": {"OfficeApi": true, "WordApi": true}
    })
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

    let pc_cap = pc_ready["capability"].clone();
    send(&mut pc, json!({"version":1,"type":"pc.start","session_id":sid,"capability":pc_cap,"request_id":"r1","status":200,"content_type":"text/event-stream"})).await;
    send(&mut pc, json!({"version":1,"type":"pc.chunk","session_id":sid,"capability":pc_cap,"request_id":"r1","sequence":0,"data":"bGF0ZQ=="})).await;
    send(&mut pc, json!({"version":1,"type":"pc.done","session_id":sid,"capability":pc_cap,"request_id":"r1"})).await;
    send(&mut pc, json!({"version":1,"type":"pc.error","session_id":sid,"capability":pc_cap,"request_id":"r1","code":"cancelled"})).await;

    send(&mut office, json!({"version":1,"type":"office.request","session_id":sid,"capability":cap,"request_id":"r2","body":{"messages":[]}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "r2");
    send(&mut pc, json!({"version":1,"type":"pc.start","session_id":sid,"capability":pc_cap,"request_id":"r2","status":200,"content_type":"text/event-stream"})).await;
    assert_eq!(recv(&mut office).await["type"], "relay.start");
    send(&mut pc, json!({"version":1,"type":"pc.chunk","session_id":sid,"capability":pc_cap,"request_id":"r2","sequence":0,"data":"b2s="})).await;
    assert_eq!(recv(&mut office).await["data"], "b2s=");
    send(&mut pc, json!({"version":1,"type":"pc.done","session_id":sid,"capability":pc_cap,"request_id":"r2"})).await;
    assert_eq!(recv(&mut office).await["type"], "relay.done");
    // Response stream cleanup can race behind pc.done/pc.error. A late cancel for
    // a request that Relay already made terminal must be idempotent.
    send(&mut office, json!({"version":1,"type":"office.cancel","session_id":sid,"capability":cap,"request_id":"r2"})).await;

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
        json!({"version":2,"type":"office.create","host":"Word","capabilities":["agent.v1","web-search.v1","web-fetch.v1","future-capability.v9"]}),
    )
    .await;
    let created = recv(&mut office).await;
    assert_eq!(created["version"], 2);
    let code = created["verification_code"].as_str().unwrap();

    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":2,"type":"pc.negotiate","verification_code":code,"capabilities":["agent.v1","web-search.v1","image-search.v1","future-capability.v9"]}),
    )
    .await;
    let negotiated = recv(&mut pc).await;
    assert_eq!(negotiated["type"], "pc.negotiated");
    assert_eq!(negotiated["pairing_version"], 2);
    assert_eq!(
        negotiated["capabilities"],
        json!(["agent.v1", "web-search.v1"])
    );
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
    send(&mut pc, json!({"version":2,"type":"pc.start","session_id":office_ready["session_id"],"capability":pc_ready["capability"],"request_id":"web_request_2","status":200,"content_type":"application/json"})).await;
    assert_eq!(recv(&mut office).await["version"], 2);
    send(&mut pc, json!({"version":2,"type":"pc.chunk","session_id":office_ready["session_id"],"capability":pc_ready["capability"],"request_id":"web_request_2","sequence":0,"data":"e30="})).await;
    assert_eq!(recv(&mut office).await["version"], 2);
    send(&mut pc, json!({"version":2,"type":"pc.done","session_id":office_ready["session_id"],"capability":pc_ready["capability"],"request_id":"web_request_2"})).await;
    assert_eq!(recv(&mut office).await["version"], 2);
    pc.close(None).await.unwrap();
    let revoked = recv(&mut office).await;
    assert_eq!(revoked["version"], 2);
    assert_eq!(revoked["code"], "session_revoked");
}

#[tokio::test]
async fn three_valid_v2_negotiations_and_claims_share_one_subject_budget() {
    let url = server().await;
    let mut office = socket(&url, ORIGIN).await;
    let mut pcs = Vec::new();

    for host in ["Word", "Excel", "PowerPoint"] {
        send(
            &mut office,
            json!({"version":2,"type":"office.create","host":host,"capabilities":["agent.v1"]}),
        )
        .await;
        let created = recv(&mut office).await;
        let code = created["verification_code"].clone();
        let mut pc = pc_socket(&url).await;
        send(
            &mut pc,
            json!({"version":2,"type":"pc.negotiate","verification_code":code,"capabilities":["agent.v1"]}),
        )
        .await;
        let negotiated = recv(&mut pc).await;
        assert_eq!(negotiated["type"], "pc.negotiated", "{negotiated}");
        send(
            &mut pc,
            json!({"version":2,"type":"pc.claim","verification_code":code,"capabilities":["agent.v1"]}),
        )
        .await;
        assert_eq!(recv(&mut pc).await["type"], "pc.claimed");
        pcs.push(pc);
    }
}

#[tokio::test]
async fn negotiated_claims_do_not_double_consume_the_global_claim_budget() {
    let url = server_with_limits(None, Some(3)).await;
    let mut office = socket(&url, ORIGIN).await;
    let mut pcs = Vec::new();

    for host in ["Word", "Excel"] {
        send(
            &mut office,
            json!({"version":2,"type":"office.create","host":host,"capabilities":["agent.v1"]}),
        )
        .await;
        let created = recv(&mut office).await;
        let code = created["verification_code"].clone();
        let mut pc = pc_socket(&url).await;
        send(
            &mut pc,
            json!({"version":2,"type":"pc.negotiate","verification_code":code,"capabilities":["agent.v1"]}),
        )
        .await;
        assert_eq!(recv(&mut pc).await["type"], "pc.negotiated");
        send(
            &mut pc,
            json!({"version":2,"type":"pc.claim","verification_code":code,"capabilities":["agent.v1"]}),
        )
        .await;
        assert_eq!(recv(&mut pc).await["type"], "pc.claimed");
        pcs.push(pc);
    }

    let mut invalid = pc_socket(&url).await;
    send(
        &mut invalid,
        json!({"version":1,"type":"pc.claim","verification_code":"999999"}),
    )
    .await;
    assert_eq!(recv(&mut invalid).await["code"], "invalid_code");

    let mut limited = pc_socket(&url).await;
    send(
        &mut limited,
        json!({"version":1,"type":"pc.claim","verification_code":"999999"}),
    )
    .await;
    assert_eq!(recv(&mut limited).await["code"], "claim_rate_limited");
}

#[tokio::test]
async fn invalid_v2_codes_remain_subject_limited_across_reconnects() {
    let url = server().await;
    for attempt in 0..6 {
        let mut pc = pc_socket(&url).await;
        send(
            &mut pc,
            json!({"version":2,"type":"pc.negotiate","verification_code":"999999","capabilities":["agent.v1"]}),
        )
        .await;
        assert_eq!(
            recv(&mut pc).await["code"],
            if attempt < 5 {
                "invalid_code"
            } else {
                "claim_limit"
            }
        );
    }
}

#[tokio::test]
async fn valid_activity_renews_idle_ttl_but_never_the_absolute_session_lifetime() {
    let url = server_with_session_ttls(Some((
        Duration::from_millis(150),
        Duration::from_millis(450),
    )))
    .await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"Word"}),
    )
    .await;
    let created = recv(&mut office).await;
    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.claim","verification_code":created["verification_code"]}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.approve","pairing_id":claimed["pairing_id"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;

    tokio::time::sleep(Duration::from_millis(100)).await;
    send(&mut office, json!({"version":1,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"renew_request_1","body":{}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "renew_request_1");
    send(&mut office, json!({"version":1,"type":"office.cancel","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"renew_request_1"})).await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");

    tokio::time::sleep(Duration::from_millis(100)).await;
    send(&mut office, json!({"version":1,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"renew_request_2","body":{}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "renew_request_2");
    send(&mut office, json!({"version":1,"type":"office.cancel","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"renew_request_2"})).await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");

    tokio::time::sleep(Duration::from_millis(270)).await;
    send(&mut office, json!({"version":1,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"past_absolute","body":{}})).await;
    let expired = recv(&mut office).await;
    assert_eq!(expired["code"], "session_expired");
    let pc_expired = recv(&mut pc).await;
    assert_eq!(pc_expired["code"], "session_expired");
    assert_eq!(pc_ready["session_id"], office_ready["session_id"]);
}

#[tokio::test]
async fn websocket_liveness_renews_an_idle_cowork_session() {
    let url = server_with_session_ttls(Some((
        Duration::from_millis(120),
        Duration::from_millis(500),
    )))
    .await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"PowerPoint"}),
    )
    .await;
    let created = recv(&mut office).await;
    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.claim","verification_code":created["verification_code"]}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.approve","pairing_id":claimed["pairing_id"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;

    tokio::time::sleep(Duration::from_millis(80)).await;
    pc.send(Message::Pong(Vec::new().into())).await.unwrap();
    tokio::time::sleep(Duration::from_millis(80)).await;
    send(&mut office, json!({"version":1,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"after_idle_heartbeat","body":{}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "after_idle_heartbeat");
    assert_eq!(pc_ready["session_id"], office_ready["session_id"]);
}

#[tokio::test]
async fn unknown_only_claim_does_not_mutate_or_notify_the_pairing() {
    let url = server().await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.create","host":"Word","capabilities":["agent.v1"]}),
    )
    .await;
    let created = recv(&mut office).await;
    let code = created["verification_code"].clone();

    for attempt in 0..5 {
        let mut attacker = pc_socket(&url).await;
        let message_type = if attempt % 2 == 0 {
            "pc.negotiate"
        } else {
            "pc.claim"
        };
        send(
            &mut attacker,
            json!({"version":2,"type":message_type,"verification_code":code,"capabilities":["future-capability.v9"]}),
        )
        .await;
        assert_eq!(
            recv(&mut attacker).await["code"],
            "capability_not_negotiated"
        );
    }
    assert!(
        tokio::time::timeout(Duration::from_millis(50), recv(&mut office))
            .await
            .is_err()
    );

    let mut legitimate = pc_socket_with_token(&url, "legitimate-test-token")
        .await
        .unwrap();
    send(
        &mut legitimate,
        json!({"version":2,"type":"pc.claim","verification_code":code,"capabilities":["agent.v1"]}),
    )
    .await;
    assert_eq!(recv(&mut legitimate).await["type"], "pc.claimed");
}

#[tokio::test]
async fn pc_response_activity_cannot_revive_an_idle_expired_session() {
    let url =
        server_with_session_ttls(Some((Duration::from_millis(100), Duration::from_secs(1)))).await;
    let mut office = socket(&url, ORIGIN).await;
    send(
        &mut office,
        json!({"version":1,"type":"office.create","host":"Excel"}),
    )
    .await;
    let created = recv(&mut office).await;
    let mut pc = pc_socket(&url).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.claim","verification_code":created["verification_code"]}),
    )
    .await;
    let claimed = recv(&mut pc).await;
    send(
        &mut pc,
        json!({"version":1,"type":"pc.approve","pairing_id":claimed["pairing_id"]}),
    )
    .await;
    let pc_ready = recv(&mut pc).await;
    let office_ready = recv(&mut office).await;
    send(&mut office, json!({"version":1,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"idle_response","body":{}})).await;
    assert_eq!(recv(&mut pc).await["request_id"], "idle_response");

    tokio::time::sleep(Duration::from_millis(150)).await;
    send(&mut pc, json!({"version":1,"type":"pc.start","session_id":pc_ready["session_id"],"capability":pc_ready["capability"],"request_id":"idle_response","status":200,"content_type":"application/json"})).await;
    let expired = recv(&mut office).await;
    assert_eq!(expired["code"], "session_expired");
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
    let invalid_frame = recv(&mut invalid).await;
    assert_eq!(invalid_frame["version"], 2);
    assert_eq!(invalid_frame["code"], "invalid_frame");
    let mut binary = socket(&url, ORIGIN).await;
    binary
        .send(Message::Binary(vec![1, 2, 3].into()))
        .await
        .unwrap();
    assert_eq!(recv(&mut binary).await["code"], "binary_not_supported");
}

#[tokio::test]
async fn accepts_capability_bound_diagnostic_without_forwarding_or_disturbing_requests() {
    let url = server().await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;

    send(
        &mut office,
        diagnostic(&office_ready, "7c24d89e-c125-43d1-baf5-017f517fe269"),
    )
    .await;
    let accepted = recv(&mut office).await;
    assert_eq!(accepted["version"], 2);
    assert_eq!(accepted["type"], "office.diagnostic.accepted");
    assert_eq!(accepted["event_id"], "7c24d89e-c125-43d1-baf5-017f517fe269");
    assert!(
        tokio::time::timeout(Duration::from_millis(50), recv(&mut pc))
            .await
            .is_err()
    );

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"after_diagnostic","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "after_diagnostic");

    send(
        &mut office,
        diagnostic(&office_ready, "b6811fa8-e737-47b8-ae6a-27881720a977"),
    )
    .await;
    assert_eq!(
        recv(&mut office).await["type"],
        "office.diagnostic.accepted"
    );
    assert!(
        tokio::time::timeout(Duration::from_millis(50), recv(&mut pc))
            .await
            .is_err()
    );
    send(
        &mut office,
        json!({"version":2,"type":"office.cancel","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"after_diagnostic"}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");
}

#[tokio::test]
async fn oversized_diagnostic_is_nonfatal_while_an_agent_request_is_active() {
    let url = server().await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;
    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"active_during_oversize","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "active_during_oversize");

    let mut oversized = diagnostic(&office_ready, "04685f50-15c0-4392-a361-b0f36a84a719");
    oversized["office_error_location"] = json!("a".repeat(17 * 1024));
    send(&mut office, oversized).await;
    assert_eq!(recv(&mut office).await["code"], "diagnostic_too_large");
    assert!(
        tokio::time::timeout(Duration::from_millis(50), recv(&mut pc))
            .await
            .is_err()
    );

    send(
        &mut office,
        json!({"version":2,"type":"office.cancel","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"active_during_oversize"}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");
}

#[tokio::test]
async fn rejects_unbound_or_non_v2_diagnostics_without_ending_the_office_connection() {
    let url = server().await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;

    let mut wrong_capability = diagnostic(&office_ready, "c7a49ea2-349c-42d8-b1db-87668d2ec8e2");
    wrong_capability["capability"] = json!("not-the-office-capability");
    send(&mut office, wrong_capability).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_capability");

    let mut wrong_session = diagnostic(&office_ready, "bdba1cbd-a72f-4858-a924-f6a06981041b");
    wrong_session["session_id"] = json!("not-the-session");
    send(&mut office, wrong_session).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_session");

    let mut old_protocol = diagnostic(&office_ready, "83c03510-6b4b-430f-afce-a98f4c5346a4");
    old_protocol["version"] = json!(1);
    send(&mut office, old_protocol).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_frame");

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"still_alive","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "still_alive");
}

#[tokio::test]
async fn rejects_unknown_sensitive_and_unbounded_diagnostic_fields() {
    let url = server().await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;

    let mut sensitive = diagnostic(&office_ready, "63878550-b3c8-46a2-a137-4473d3a2fcd3");
    sensitive["message"] = json!("secret document text");
    send(&mut office, sensitive).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_frame");

    let mut unknown_error = diagnostic(&office_ready, "e1ee1f53-b40e-4852-a208-81292f044abb");
    unknown_error["error_code"] = json!("contains_user_content");
    send(&mut office, unknown_error).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_frame");

    let mut unsafe_identifier = diagnostic(&office_ready, "0314e2c2-f916-4826-a7c4-e0ffc7f58b2d");
    unsafe_identifier["office_error_location"] = json!("Body.insertText\nsecret");
    send(&mut office, unsafe_identifier).await;
    assert_eq!(recv(&mut office).await["code"], "invalid_frame");

    let mut oversized = diagnostic(&office_ready, "85c91123-b569-42d4-9d0a-39a2df9a6d71");
    oversized["office_error_location"] = json!("a".repeat(3_900));
    send(&mut office, oversized).await;
    assert_eq!(recv(&mut office).await["code"], "diagnostic_too_large");

    assert!(
        tokio::time::timeout(Duration::from_millis(50), recv(&mut pc))
            .await
            .is_err()
    );
}

#[tokio::test]
async fn rate_limits_diagnostic_bursts_and_recovers_without_disturbing_requests() {
    let url = server_with_diagnostic_rate(10, Duration::from_millis(100)).await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;

    for index in 0..10 {
        let event_id = format!("10000000-0000-4000-8000-{index:012}");
        send(&mut office, diagnostic(&office_ready, &event_id)).await;
        assert_eq!(
            recv(&mut office).await["type"],
            "office.diagnostic.accepted"
        );
    }
    send(
        &mut office,
        diagnostic(&office_ready, "10000000-0000-4000-8000-000000000010"),
    )
    .await;
    assert_eq!(recv(&mut office).await["code"], "diagnostic_rate_limited");

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"during_rate_limit","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "during_rate_limit");
    send(
        &mut office,
        json!({"version":2,"type":"office.cancel","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"during_rate_limit"}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["type"], "relay.cancel");

    tokio::time::sleep(Duration::from_millis(110)).await;
    send(
        &mut office,
        diagnostic(&office_ready, "10000000-0000-4000-8000-000000000011"),
    )
    .await;
    assert_eq!(
        recv(&mut office).await["type"],
        "office.diagnostic.accepted"
    );
}

#[tokio::test]
async fn diagnostic_host_must_match_the_approved_pairing_host() {
    let url = server().await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;
    let mut mismatched = diagnostic(&office_ready, "72447ee2-b34f-432e-9835-ab0ae7674546");
    mismatched["host"] = json!("excel");
    send(&mut office, mismatched).await;
    assert_eq!(recv(&mut office).await["code"], "diagnostic_host_mismatch");

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"after_host_mismatch","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "after_host_mismatch");
}

#[tokio::test]
async fn caps_diagnostics_at_one_hundred_per_session_without_mutating_request_state() {
    let url = server_with_diagnostic_rate(100, Duration::from_secs(1)).await;
    let (mut office, mut pc, office_ready, _) = approved_v2_session(&url).await;

    for index in 0..100 {
        let event_id = format!("00000000-0000-4000-8000-{index:012}");
        send(&mut office, diagnostic(&office_ready, &event_id)).await;
        assert_eq!(
            recv(&mut office).await["type"],
            "office.diagnostic.accepted"
        );
    }
    send(
        &mut office,
        diagnostic(&office_ready, "00000000-0000-4000-8000-000000000100"),
    )
    .await;
    assert_eq!(recv(&mut office).await["code"], "diagnostic_limit");

    send(
        &mut office,
        json!({"version":2,"type":"office.request","session_id":office_ready["session_id"],"capability":office_ready["capability"],"request_id":"after_limit","capability_name":"agent.v1","body":{}}),
    )
    .await;
    assert_eq!(recv(&mut pc).await["request_id"], "after_limit");
}
