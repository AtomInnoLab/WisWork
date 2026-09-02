mod binding_store;

use axum::{
    Router,
    extract::{
        ConnectInfo, State,
        ws::{Message, WebSocket, WebSocketUpgrade},
    },
    http::{HeaderMap, StatusCode},
    response::IntoResponse,
    routing::get,
};
use base64::{
    Engine as _,
    engine::general_purpose::{STANDARD, URL_SAFE_NO_PAD},
};
use binding_store::{Binding, BindingStore};
use futures_util::{SinkExt, StreamExt};
use jsonwebtoken::{Algorithm, DecodingKey, Validation, decode, decode_header, jwk::JwkSet};
use p256::ecdsa::{Signature, VerifyingKey, signature::Verifier};
use rand::{Rng, distr::Alphanumeric};
use serde::Deserialize;
use serde_json::{Map, Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::{HashMap, HashSet, VecDeque},
    net::{IpAddr, SocketAddr},
    path::PathBuf,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};
use tokio::sync::{Mutex, Notify, Semaphore, mpsc};

pub use binding_store::BindingStoreError;

pub const OFFICE_ORIGIN: &str = "https://office.8-216-134-194.sslip.io";
const CONTROL_MAX: usize = 16 * 1024;
const REQUEST_MAX: usize = 256 * 1024;
const FRAME_MAX: usize = REQUEST_MAX + CONTROL_MAX;
const CHUNK_MAX: usize = 64 * 1024;
const RESPONSE_MAX: usize = 16 * 1024 * 1024;
const DIAGNOSTIC_MAX: usize = 4 * 1024;
const DIAGNOSTIC_SESSION_MAX: u16 = 100;
const PROTOCOL_V2: u64 = 2;
const PAIRING_RESUME: &str = "pairing-resume.v1";
const RESUME_CHALLENGE_MAX: Duration = Duration::from_secs(30);
const MAX_TRACKED_RESUME_IPS: usize = 10_000;
const SUPPORTED_CAPABILITIES: &[&str] = &[
    "agent.v1",
    "web-search.v1",
    "web-fetch.v1",
    "image-search.v1",
];

#[derive(Clone)]
pub struct Config {
    pub pairing_ttl: Duration,
    pub session_ttl: Duration,
    pub session_max_ttl: Duration,
    pub request_ttl: Duration,
    pub max_claim_attempts: u8,
    pub max_global_claims: u32,
    pub max_global_resume_attempts: u32,
    pub diagnostic_window: Duration,
    pub max_diagnostics_per_window: u8,
    pub auth_url: String,
    pub jwks_url: String,
    pub issuer: String,
    pub audience: String,
    pub binding_database: Option<PathBuf>,
    pub pairing_resume_enabled: bool,
    pub resume_challenge_ttl: Duration,
}
impl Default for Config {
    fn default() -> Self {
        Self {
            pairing_ttl: Duration::from_secs(120),
            session_ttl: Duration::from_secs(1800),
            session_max_ttl: Duration::from_secs(8 * 60 * 60),
            request_ttl: Duration::from_secs(300),
            max_claim_attempts: 5,
            max_global_claims: 1_000,
            max_global_resume_attempts: 1_000,
            diagnostic_window: Duration::from_secs(1),
            max_diagnostics_per_window: 10,
            auth_url: "https://auth.wispaper.ai/oidc/me".into(),
            jwks_url: "https://auth.wispaper.ai/oidc/jwks".into(),
            issuer: "https://auth.wispaper.ai/oidc".into(),
            audience: "i9au2rbqzktme4runr9gy".into(),
            binding_database: None,
            pairing_resume_enabled: true,
            resume_challenge_ttl: RESUME_CHALLENGE_MAX,
        }
    }
}

#[derive(Clone)]
struct App {
    inner: Arc<Inner>,
}
struct Inner {
    state: Mutex<Store>,
    next: AtomicU64,
    config: Config,
    http: reqwest::Client,
    auth_slots: Arc<Semaphore>,
    bindings: BindingStore,
    #[cfg(test)]
    fail_approval_delivery: std::sync::atomic::AtomicBool,
}
#[derive(Clone)]
struct Tx {
    sender: mpsc::Sender<Message>,
    failed: Arc<Notify>,
}
struct Pairing {
    version: u64,
    id: String,
    code: String,
    host: String,
    office: u64,
    office_tx: Tx,
    pc: Option<(u64, Tx)>,
    expires: Instant,
    attempts: u8,
    requested_capabilities: Vec<String>,
    negotiated_capabilities: Vec<String>,
    negotiated_subject: Option<[u8; 32]>,
    features: Vec<String>,
    negotiated_features: Vec<String>,
    binding_public_key: Option<[u8; 65]>,
    pc_subject: Option<[u8; 32]>,
    pending_binding: Option<PendingBinding>,
}
#[derive(Clone)]
struct PendingBinding {
    id: String,
    pc: u64,
    pc_tx: Tx,
    subject: [u8; 32],
    capabilities: Vec<String>,
    phase: PendingBindingPhase,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum PendingBindingPhase {
    Offered,
    CommitSent,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum BindingResult {
    Ready,
    Abort,
    Committed,
}
struct Active {
    id: String,
    sequence: u64,
    bytes: usize,
    deadline: Instant,
    started: bool,
    pending_tool: Option<PendingTool>,
    used_tool_calls: VecDeque<String>,
}
struct PendingTool {
    turn_id: String,
    call_id: String,
    generation: u64,
}
struct Session {
    version: u64,
    host: String,
    office: u64,
    office_tx: Tx,
    pc: u64,
    pc_tx: Tx,
    office_cap: String,
    pc_cap: String,
    expires: Instant,
    absolute_expires: Instant,
    active: Option<Active>,
    used_requests: VecDeque<String>,
    capabilities: Vec<String>,
    diagnostics: u16,
    diagnostic_window_started: Instant,
    diagnostic_window_count: u8,
    binding_id: Option<String>,
}
struct ResumeChallenge {
    binding: Binding,
    challenge: String,
    expires: Instant,
}
struct ExpiredResumeChallenge {
    binding_hash: [u8; 32],
    challenge_hash: [u8; 32],
    expires: Instant,
}
struct CompletedBindingOffer {
    office: u64,
    binding_id: String,
    expires: Instant,
}
struct OfficeResume {
    binding: Binding,
    office: u64,
    office_tx: Tx,
    expires: Instant,
}
struct PcResume {
    binding_id: String,
    pc: u64,
    pc_tx: Tx,
    subject: [u8; 32],
    capabilities: Vec<String>,
    expires: Instant,
}
#[derive(Default)]
struct Store {
    pairings: HashMap<String, Pairing>,
    codes: HashMap<String, String>,
    sessions: HashMap<String, Session>,
    claim_attempts: HashMap<[u8; 32], (Instant, u8)>,
    create_attempts: HashMap<IpAddr, (Instant, u8)>,
    global_claims: (Option<Instant>, u32),
    global_creates: (Option<Instant>, u32),
    connection_pairings: HashMap<u64, u8>,
    connection_sessions: HashMap<u64, HashSet<String>>,
    preauth_attempts: HashMap<IpAddr, (Instant, u8)>,
    global_preauth: (Option<Instant>, u32),
    resume_challenges: HashMap<u64, ResumeChallenge>,
    expired_resume_challenges: HashMap<u64, ExpiredResumeChallenge>,
    completed_binding_offers: HashMap<String, CompletedBindingOffer>,
    office_resumes: HashMap<u64, OfficeResume>,
    pc_resumes: HashMap<u64, PcResume>,
    connection_resume_attempts: HashMap<u64, u8>,
    resume_attempts: HashMap<IpAddr, (Instant, u8)>,
    global_resume_attempts: (Option<Instant>, u32),
    denied_bindings: HashSet<String>,
    pending_revocations: HashMap<String, [u8; 32]>,
}

pub fn app(config: Config) -> Router {
    try_app(config).expect("binding database initialization failed")
}

pub fn try_app(config: Config) -> Result<Router, BindingStoreError> {
    let bindings = if config.pairing_resume_enabled {
        BindingStore::open(config.binding_database.as_deref())?
    } else {
        BindingStore::disabled()
    };
    let state = App {
        inner: Arc::new(Inner {
            state: Mutex::new(Store::default()),
            next: AtomicU64::new(1),
            config,
            http: reqwest::Client::builder()
                .no_proxy()
                .timeout(Duration::from_secs(3))
                .redirect(reqwest::redirect::Policy::none())
                .build()
                .expect("fixed HTTP client"),
            auth_slots: Arc::new(Semaphore::new(32)),
            bindings,
            #[cfg(test)]
            fail_approval_delivery: std::sync::atomic::AtomicBool::new(false),
        }),
    };
    drop(spawn_sweeper(&state));
    Ok(Router::new()
        .route("/office-relay", get(upgrade))
        .route("/office-relay/health", get(|| async { "ok" }))
        .with_state(state))
}

fn spawn_sweeper(state: &App) -> tokio::task::JoinHandle<()> {
    let sweeper = Arc::downgrade(&state.inner);
    tokio::spawn(async move {
        let mut interval = tokio::time::interval(Duration::from_secs(1));
        loop {
            interval.tick().await;
            let Some(inner) = sweeper.upgrade() else {
                break;
            };
            let mut store = inner.state.lock().await;
            expire(&inner, &mut store);
        }
    })
}

async fn upgrade(
    State(app): State<App>,
    ConnectInfo(peer): ConnectInfo<SocketAddr>,
    headers: HeaderMap,
    ws: WebSocketUpgrade,
) -> impl IntoResponse {
    let origin = headers
        .get("origin")
        .and_then(|v| v.to_str().ok())
        .map(str::to_owned);
    if origin.as_deref().is_some_and(|v| v != OFFICE_ORIGIN) {
        return StatusCode::FORBIDDEN.into_response();
    }
    let client_ip = match headers.get("x-real-ip").and_then(|v| v.to_str().ok()) {
        Some(value) if peer.ip().is_loopback() => match value.parse() {
            Ok(ip) => ip,
            Err(_) => return StatusCode::FORBIDDEN.into_response(),
        },
        Some(_) => return StatusCode::FORBIDDEN.into_response(),
        None => peer.ip(),
    };
    let authorization = headers.get("authorization").and_then(|v| v.to_str().ok());
    let subject = if origin.is_some() {
        if authorization.is_some() {
            return StatusCode::FORBIDDEN.into_response();
        }
        None
    } else {
        if !allow_preauth(&app, client_ip).await {
            return StatusCode::TOO_MANY_REQUESTS.into_response();
        }
        let Ok(_permit) = app.inner.auth_slots.clone().try_acquire_owned() else {
            return StatusCode::SERVICE_UNAVAILABLE.into_response();
        };
        let Some(token) = authorization
            .and_then(|v| v.strip_prefix("Bearer "))
            .filter(|v| !v.is_empty() && v.len() <= 4096)
        else {
            return StatusCode::UNAUTHORIZED.into_response();
        };
        let Some(subject) = authenticate_pc(&app, token).await else {
            return StatusCode::UNAUTHORIZED.into_response();
        };
        Some(subject)
    };
    ws.max_message_size(FRAME_MAX)
        .on_upgrade(move |socket| connection(app, socket, origin, client_ip, subject))
        .into_response()
}

async fn allow_preauth(app: &App, ip: IpAddr) -> bool {
    let mut store = app.inner.state.lock().await;
    let ttl = app.inner.config.pairing_ttl;
    if store
        .global_preauth
        .0
        .is_none_or(|start| start.elapsed() > ttl)
    {
        store.global_preauth = (Some(Instant::now()), 0);
    }
    store.global_preauth.1 = store.global_preauth.1.saturating_add(1);
    if store.global_preauth.1 > 1_000
        || (store.preauth_attempts.len() >= 10_000 && !store.preauth_attempts.contains_key(&ip))
    {
        return false;
    }
    let entry = store
        .preauth_attempts
        .entry(ip)
        .or_insert((Instant::now(), 0));
    if entry.0.elapsed() > ttl {
        *entry = (Instant::now(), 0);
    }
    entry.1 = entry.1.saturating_add(1);
    entry.1 <= 20
}

async fn authenticate_pc(app: &App, token: &str) -> Option<[u8; 32]> {
    if token.split('.').count() == 3
        && let Some(subject) = authenticate_id_token(app, token).await
    {
        return Some(subject);
    }
    let Ok(url) = reqwest::Url::parse(&app.inner.config.auth_url) else {
        return None;
    };
    let allowed = url.as_str() == "https://auth.wispaper.ai/oidc/me"
        || (url.scheme() == "http"
            && url.host_str().is_some_and(|h| h == "127.0.0.1")
            && url.path() == "/oidc/me");
    if !allowed {
        return None;
    }
    let Ok(mut response) = app.inner.http.get(url).bearer_auth(token).send().await else {
        return None;
    };
    if !response.status().is_success()
        || response
            .content_length()
            .is_some_and(|n| n > CONTROL_MAX as u64)
    {
        return None;
    }
    let mut body = Vec::new();
    loop {
        match response.chunk().await {
            Ok(Some(chunk)) if body.len() + chunk.len() <= CONTROL_MAX => {
                body.extend_from_slice(&chunk)
            }
            Ok(None) => break,
            _ => return None,
        }
    }
    let value: Value = serde_json::from_slice(&body).ok()?;
    let sub = value.as_object()?.get("sub")?.as_str()?;
    if sub.is_empty() || sub.len() > 512 {
        return None;
    }
    Some(Sha256::digest(sub.as_bytes()).into())
}

#[derive(Deserialize)]
struct IdTokenClaims {
    sub: String,
}

async fn authenticate_id_token(app: &App, token: &str) -> Option<[u8; 32]> {
    let header = decode_header(token).ok()?;
    if header.alg != Algorithm::ES384 {
        return None;
    }
    let kid = header
        .kid
        .as_deref()
        .filter(|value| !value.is_empty() && value.len() <= 128)?;
    let url = reqwest::Url::parse(&app.inner.config.jwks_url).ok()?;
    let allowed = url.as_str() == "https://auth.wispaper.ai/oidc/jwks"
        || (url.scheme() == "http"
            && url.host_str().is_some_and(|host| host == "127.0.0.1")
            && url.path() == "/oidc/jwks");
    if !allowed {
        return None;
    }
    let mut response = app.inner.http.get(url).send().await.ok()?;
    if !response.status().is_success()
        || response
            .content_length()
            .is_some_and(|size| size > CONTROL_MAX as u64)
    {
        return None;
    }
    let mut body = Vec::new();
    loop {
        match response.chunk().await {
            Ok(Some(chunk)) if body.len() + chunk.len() <= CONTROL_MAX => {
                body.extend_from_slice(&chunk)
            }
            Ok(None) => break,
            _ => return None,
        }
    }
    let set: JwkSet = serde_json::from_slice(&body).ok()?;
    let key = set.find(kid)?;
    let decoding_key = DecodingKey::from_jwk(key).ok()?;
    let mut validation = Validation::new(Algorithm::ES384);
    validation.set_issuer(&[app.inner.config.issuer.as_str()]);
    validation.set_audience(&[app.inner.config.audience.as_str()]);
    let claims = decode::<IdTokenClaims>(token, &decoding_key, &validation)
        .ok()?
        .claims;
    if claims.sub.is_empty() || claims.sub.len() > 512 {
        return None;
    }
    Some(Sha256::digest(claims.sub.as_bytes()).into())
}

async fn connection(
    app: App,
    socket: WebSocket,
    origin: Option<String>,
    peer: IpAddr,
    subject: Option<[u8; 32]>,
) {
    let id = app.inner.next.fetch_add(1, Ordering::Relaxed);
    let (mut sink, mut stream) = socket.split();
    let (sender, mut rx) = mpsc::channel(64);
    let tx = Tx {
        sender,
        failed: Arc::new(Notify::new()),
    };
    let writer_failed = tx.failed.clone();
    let writer = tokio::spawn(async move {
        while let Some(message) = rx.recv().await {
            if !matches!(
                tokio::time::timeout(Duration::from_secs(10), sink.send(message)).await,
                Ok(Ok(()))
            ) {
                writer_failed.notify_one();
                break;
            }
        }
    });
    let mut heartbeat = tokio::time::interval(Duration::from_secs(20));
    let renewal_interval = app
        .inner
        .config
        .session_ttl
        .checked_div(3)
        .unwrap_or(Duration::from_millis(1))
        .max(Duration::from_millis(1))
        .min(Duration::from_secs(60));
    let mut last_lease_renewal = Instant::now();
    heartbeat.tick().await;
    loop {
        let message = tokio::select! {
            _ = tx.failed.notified() => break,
            _ = heartbeat.tick() => { if tx.sender.try_send(Message::Ping(Vec::new().into())).is_err() { break; } continue; },
            value = tokio::time::timeout(Duration::from_secs(60), stream.next()) => match value { Ok(value) => value, Err(_) => break },
        };
        let Some(Ok(message)) = message else { break };
        match message {
            Message::Text(text) if text.len() <= FRAME_MAX => {
                if let Err(code) = process(
                    &app,
                    id,
                    &tx,
                    origin.as_deref(),
                    peer,
                    subject,
                    text.as_str(),
                )
                .await
                {
                    eprintln!(
                        "{}",
                        protocol_error_log(id, origin.is_some(), text.as_str(), code)
                    );
                    error_for_text(&tx, text.as_str(), code);
                    break;
                }
            }
            Message::Ping(data) => {
                tx.sender
                    .try_send(Message::Pong(data))
                    .unwrap_or_else(|_| tx.failed.notify_one());
                if last_lease_renewal.elapsed() >= renewal_interval {
                    renew_connection_sessions(&app, id).await;
                    last_lease_renewal = Instant::now();
                }
            }
            Message::Pong(_) => {
                if last_lease_renewal.elapsed() >= renewal_interval {
                    renew_connection_sessions(&app, id).await;
                    last_lease_renewal = Instant::now();
                }
            }
            Message::Close(_) => break,
            Message::Binary(_) => {
                error(&tx, "binary_not_supported");
                break;
            }
            Message::Text(_) => {
                error(&tx, "frame_too_large");
                break;
            }
        }
    }
    cleanup(&app, id).await;
    drop(tx);
    let _ = writer.await;
}

fn object(text: &str, max: usize) -> Result<Map<String, Value>, &'static str> {
    if text.len() > max {
        return Err("frame_too_large");
    }
    serde_json::from_str::<Value>(text)
        .ok()
        .and_then(|v| v.as_object().cloned())
        .ok_or("invalid_frame")
}
fn exact(map: &Map<String, Value>, keys: &[&str]) -> bool {
    map.len() == keys.len() && keys.iter().all(|k| map.contains_key(*k))
}
fn string<'a>(m: &'a Map<String, Value>, k: &str) -> Result<&'a str, &'static str> {
    m.get(k).and_then(Value::as_str).ok_or("invalid_frame")
}
fn token() -> String {
    rand::rng()
        .sample_iter(&Alphanumeric)
        .take(43)
        .map(char::from)
        .collect()
}
fn send(tx: &Tx, value: Value) {
    if try_send(tx, value).is_err() {
        tx.failed.notify_one();
    }
}
fn try_send(tx: &Tx, value: Value) -> Result<(), ()> {
    tx.sender
        .try_send(Message::Text(value.to_string().into()))
        .map_err(|_| ())
}
fn error(tx: &Tx, code: &str) {
    send(tx, json!({"version":1,"type":"relay.error","code":code}));
}
fn error_for_text(tx: &Tx, text: &str, code: &str) {
    let frame_version = serde_json::from_str::<Value>(text)
        .ok()
        .and_then(|value| value.get("version").and_then(Value::as_u64))
        .filter(|value| *value == PROTOCOL_V2)
        .unwrap_or(1);
    versioned_error(tx, frame_version, code);
}
fn protocol_error_log(conn: u64, office: bool, text: &str, code: &str) -> Value {
    let frame_type = serde_json::from_str::<Value>(text)
        .ok()
        .and_then(|value| value.get("type").and_then(Value::as_str).map(str::to_owned))
        .filter(|value| {
            !value.is_empty()
                && value.len() <= 64
                && value.bytes().all(|byte| {
                    byte.is_ascii_lowercase()
                        || byte.is_ascii_digit()
                        || matches!(byte, b'.' | b'-' | b'_')
                })
        })
        .unwrap_or_else(|| "unknown".to_owned());
    json!({
        "event": "relay_protocol_error",
        "connection_id": conn,
        "role": if office { "office" } else { "pc" },
        "frame_type": frame_type,
        "error_code": code,
    })
}
fn versioned_error(tx: &Tx, version: u64, code: &str) {
    send(
        tx,
        json!({"version":version,"type":"relay.error","code":code}),
    );
}
fn diagnostic_response(tx: &Tx, value: Value) {
    // Diagnostics are optional observability. A saturated client queue must never
    // revoke or slow the Agent session merely because an ACK cannot be delivered.
    let _ = try_send(tx, value);
}
fn renew_session(session: &mut Session, idle_ttl: Duration) {
    session.expires = (Instant::now() + idle_ttl).min(session.absolute_expires);
}
async fn renew_connection_sessions(app: &App, conn: u64) {
    let mut store = app.inner.state.lock().await;
    let idle_ttl = app.inner.config.session_ttl;
    let now = Instant::now();
    let session_ids = store
        .connection_sessions
        .get(&conn)
        .cloned()
        .unwrap_or_default();
    for session_id in session_ids {
        if let Some(session) = store.sessions.get_mut(&session_id)
            && session.expires > now
        {
            renew_session(session, idle_ttl);
        }
    }
}
fn known_inactive_request(session: &Session, request_id: &str) -> bool {
    session.active.as_ref().map(|active| active.id.as_str()) != Some(request_id)
        && session.used_requests.iter().any(|used| used == request_id)
}
fn valid_base(m: &Map<String, Value>) -> bool {
    matches!(
        m.get("version").and_then(Value::as_u64),
        Some(1 | PROTOCOL_V2)
    ) && m.get("type").and_then(Value::as_str).is_some()
}

fn version(m: &Map<String, Value>) -> Result<u64, &'static str> {
    m.get("version")
        .and_then(Value::as_u64)
        .filter(|value| *value == 1 || *value == PROTOCOL_V2)
        .ok_or("invalid_frame")
}

fn capabilities(m: &Map<String, Value>) -> Result<Vec<String>, &'static str> {
    let values = m
        .get("capabilities")
        .and_then(Value::as_array)
        .ok_or("invalid_frame")?;
    if values.is_empty() || values.len() > 16 {
        return Err("invalid_frame");
    }
    let mut result = Vec::with_capacity(values.len());
    let mut seen = HashSet::with_capacity(values.len());
    for value in values {
        let name = value.as_str().ok_or("invalid_frame")?;
        if name.is_empty()
            || name.len() > 64
            || !name.bytes().all(|byte| {
                byte.is_ascii_lowercase() || byte.is_ascii_digit() || matches!(byte, b'.' | b'-')
            })
            || !seen.insert(name)
        {
            return Err("invalid_frame");
        }
        if SUPPORTED_CAPABILITIES.contains(&name) {
            result.push(name.to_owned());
        }
    }
    Ok(result)
}

fn resume_capabilities(m: &Map<String, Value>) -> Result<Vec<String>, &'static str> {
    let values = m
        .get("capabilities")
        .and_then(Value::as_array)
        .ok_or("invalid_frame")?;
    if values.is_empty() || values.len() > SUPPORTED_CAPABILITIES.len() {
        return Err("invalid_frame");
    }
    let mut result = Vec::with_capacity(values.len());
    let mut seen = HashSet::with_capacity(values.len());
    for value in values {
        let name = value.as_str().ok_or("invalid_frame")?;
        if !SUPPORTED_CAPABILITIES.contains(&name) || !seen.insert(name) {
            return Err("invalid_frame");
        }
        result.push(name.to_owned());
    }
    Ok(result)
}

fn features(m: &Map<String, Value>, enabled: bool) -> Result<Vec<String>, &'static str> {
    let values = m
        .get("features")
        .and_then(Value::as_array)
        .ok_or("invalid_frame")?;
    if values.len() > 1 {
        return Err("invalid_frame");
    }
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for value in values {
        let name = value.as_str().ok_or("invalid_frame")?;
        if name != PAIRING_RESUME || !seen.insert(name) {
            return Err("invalid_frame");
        }
        if enabled && name == PAIRING_RESUME {
            result.push(name.to_owned());
        }
    }
    Ok(result)
}

fn binding_public_key(m: &Map<String, Value>) -> Result<[u8; 65], &'static str> {
    let encoded = string(m, "binding_public_key")?;
    let decoded = URL_SAFE_NO_PAD
        .decode(encoded)
        .map_err(|_| "invalid_frame")?;
    let key: [u8; 65] = decoded.try_into().map_err(|_| "invalid_frame")?;
    if key[0] != 4 || VerifyingKey::from_sec1_bytes(&key).is_err() {
        return Err("invalid_frame");
    }
    Ok(key)
}

async fn process(
    app: &App,
    conn: u64,
    tx: &Tx,
    origin: Option<&str>,
    peer: IpAddr,
    subject: Option<[u8; 32]>,
    text: &str,
) -> Result<(), &'static str> {
    let map = object(text, FRAME_MAX)?;
    if !valid_base(&map) {
        return Err("invalid_frame");
    }
    let kind = string(&map, "type")?;
    if text.len() > CONTROL_MAX
        && !matches!(
            kind,
            "office.request"
                | "office.diagnostic"
                | "pc.chunk"
                | "pc.tool_call"
                | "office.tool_result"
        )
    {
        return Err("frame_too_large");
    }
    if origin.is_some() && !kind.starts_with("office.") {
        return Err("role_not_allowed");
    }
    if origin.is_none() && !kind.starts_with("pc.") {
        return Err("role_not_allowed");
    }
    match kind {
        "office.create" => create(app, conn, tx, origin, peer, map).await,
        "office.resume" => office_resume(app, conn, tx, origin, peer, map).await,
        "office.prove" if app.inner.config.pairing_resume_enabled => {
            office_prove(app, conn, tx, map).await
        }
        "office.binding_ready" if app.inner.config.pairing_resume_enabled => {
            office_binding_result(app, conn, map, BindingResult::Ready).await
        }
        "office.binding_abort" if app.inner.config.pairing_resume_enabled => {
            office_binding_result(app, conn, map, BindingResult::Abort).await
        }
        "office.binding_committed" if app.inner.config.pairing_resume_enabled => {
            office_binding_result(app, conn, map, BindingResult::Committed).await
        }
        "pc.negotiate" => negotiate(app, tx, subject.ok_or("auth_required")?, map).await,
        "pc.claim" => claim(app, conn, tx, subject.ok_or("auth_required")?, map).await,
        "pc.approve" => approve(app, conn, tx, map).await,
        "pc.reject" => reject(app, conn, map).await,
        "pc.resume" => pc_resume(app, conn, tx, subject.ok_or("auth_required")?, map).await,
        "pc.revoke_binding" if app.inner.config.pairing_resume_enabled => {
            revoke_binding(app, conn, tx, subject.ok_or("auth_required")?, map).await
        }
        "office.request" => request(app, conn, map).await,
        "office.cancel" => cancel(app, conn, map).await,
        "office.tool_result" => tool_result(app, conn, map).await,
        "office.diagnostic" => {
            if let Err(code) = diagnostic(app, conn, tx, map, text.len()).await {
                diagnostic_response(
                    tx,
                    json!({"version":PROTOCOL_V2,"type":"relay.error","code":code}),
                );
            }
            Ok(())
        }
        "pc.chunk" => chunk(app, conn, map).await,
        "pc.start" => start(app, conn, map).await,
        "pc.session_state" => session_state(app, conn, map).await,
        "pc.tool_call" => tool_call(app, conn, map).await,
        "pc.done" => done(app, conn, map).await,
        "pc.error" => pc_error(app, conn, map).await,
        _ => Err("unknown_type"),
    }
}

async fn negotiate(
    app: &App,
    tx: &Tx,
    subject: [u8; 32],
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    let enhanced = m.contains_key("features");
    let expected: &[&str] = if enhanced {
        &[
            "version",
            "type",
            "verification_code",
            "capabilities",
            "features",
        ]
    } else {
        &["version", "type", "verification_code", "capabilities"]
    };
    if version(&m)? != PROTOCOL_V2 || !exact(&m, expected) {
        return Err("invalid_frame");
    }
    let offered = capabilities(&m)?;
    let offered_features = if enhanced {
        features(&m, app.inner.config.pairing_resume_enabled)?
    } else {
        Vec::new()
    };
    let code = string(&m, "verification_code")?;
    if code.len() != 6 || !code.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err("invalid_code");
    }
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    if store
        .global_claims
        .0
        .is_none_or(|start| start.elapsed() > app.inner.config.pairing_ttl)
    {
        store.global_claims = (Some(Instant::now()), 0);
    }
    store.global_claims.1 = store.global_claims.1.saturating_add(1);
    if store.global_claims.1 > app.inner.config.max_global_claims {
        return Err("claim_rate_limited");
    }
    if store.claim_attempts.len() >= 10_000 && !store.claim_attempts.contains_key(&subject) {
        return Err("relay_busy");
    }
    let attempts = store
        .claim_attempts
        .entry(subject)
        .or_insert((Instant::now(), 0));
    if attempts.0.elapsed() > app.inner.config.pairing_ttl {
        *attempts = (Instant::now(), 0);
    }
    attempts.1 = attempts.1.saturating_add(1);
    if attempts.1 > app.inner.config.max_claim_attempts {
        return Err("claim_limit");
    }
    let id = store.codes.get(code).cloned().ok_or("invalid_code")?;
    let maximum = app.inner.config.max_claim_attempts;
    let pairing = store.pairings.get_mut(&id).ok_or("invalid_code")?;
    if pairing.pc.is_some() {
        return Err("already_claimed");
    }
    let negotiated: Vec<_> = pairing
        .requested_capabilities
        .iter()
        .filter(|name| offered.contains(name))
        .cloned()
        .collect();
    if negotiated.is_empty() {
        return Err("capability_not_negotiated");
    }
    pairing.attempts = pairing.attempts.saturating_add(1);
    if pairing.attempts > maximum {
        return Err("claim_limit");
    }
    pairing.negotiated_capabilities.clone_from(&negotiated);
    pairing.negotiated_features = pairing
        .features
        .iter()
        .filter(|name| offered_features.contains(name))
        .cloned()
        .collect();
    pairing.negotiated_subject = Some(subject);
    send(
        tx,
        if enhanced {
            json!({"version":PROTOCOL_V2,"type":"pc.negotiated","pairing_version":pairing.version,"capabilities":negotiated,"features":pairing.negotiated_features})
        } else {
            json!({"version":PROTOCOL_V2,"type":"pc.negotiated","pairing_version":pairing.version,"capabilities":negotiated})
        },
    );
    Ok(())
}

async fn create(
    app: &App,
    conn: u64,
    tx: &Tx,
    origin: Option<&str>,
    peer: IpAddr,
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    let enhanced = protocol == PROTOCOL_V2 && m.contains_key("features");
    let expected: &[&str] = if enhanced {
        &[
            "version",
            "type",
            "host",
            "capabilities",
            "features",
            "binding_public_key",
        ]
    } else if protocol == PROTOCOL_V2 {
        &["version", "type", "host", "capabilities"]
    } else {
        &["version", "type", "host"]
    };
    if !exact(&m, expected) || origin != Some(OFFICE_ORIGIN) {
        return Err("invalid_frame");
    }
    let requested_capabilities = if protocol == PROTOCOL_V2 {
        capabilities(&m)?
    } else {
        vec!["agent.v1".to_owned()]
    };
    let requested_features = if enhanced {
        features(&m, app.inner.config.pairing_resume_enabled)?
    } else {
        Vec::new()
    };
    let public_key = if enhanced {
        Some(binding_public_key(&m)?)
    } else {
        None
    };
    let host = string(&m, "host")?;
    if !matches!(host, "Word" | "Excel" | "PowerPoint") {
        return Err("unsupported_host");
    }
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    let per_connection = store.connection_pairings.entry(conn).or_default();
    if *per_connection >= 4 {
        return Err("pairing_limit");
    }
    *per_connection += 1;
    if store.create_attempts.len() >= 10_000 && !store.create_attempts.contains_key(&peer) {
        return Err("relay_busy");
    }
    let per_ip = store
        .create_attempts
        .entry(peer)
        .or_insert((Instant::now(), 0));
    if per_ip.0.elapsed() > app.inner.config.pairing_ttl {
        *per_ip = (Instant::now(), 0);
    }
    per_ip.1 = per_ip.1.saturating_add(1);
    if per_ip.1 > 10 {
        return Err("create_rate_limited");
    }
    if store
        .global_creates
        .0
        .is_none_or(|start| start.elapsed() > app.inner.config.pairing_ttl)
    {
        store.global_creates = (Some(Instant::now()), 0);
    }
    store.global_creates.1 = store.global_creates.1.saturating_add(1);
    if store.global_creates.1 > 1_000 {
        return Err("relay_busy");
    }
    if store.pairings.len() >= 10_000 {
        return Err("relay_busy");
    }
    let id = token();
    let mut rng = rand::rng();
    let mut code = format!("{:06}", rng.random_range(0..1_000_000));
    while store.codes.contains_key(&code) {
        code = format!("{:06}", rng.random_range(0..1_000_000));
    }
    let pairing = Pairing {
        version: protocol,
        id: id.clone(),
        code: code.clone(),
        host: host.into(),
        office: conn,
        office_tx: tx.clone(),
        pc: None,
        expires: Instant::now() + app.inner.config.pairing_ttl,
        attempts: 0,
        requested_capabilities,
        negotiated_capabilities: Vec::new(),
        negotiated_subject: None,
        features: requested_features.clone(),
        negotiated_features: Vec::new(),
        binding_public_key: public_key,
        pc_subject: None,
        pending_binding: None,
    };
    store.codes.insert(code.clone(), id.clone());
    store.pairings.insert(id.clone(), pairing);
    send(
        tx,
        if enhanced {
            json!({"version":protocol,"type":"office.created","pairing_id":id,"verification_code":code,"expires_in":app.inner.config.pairing_ttl.as_secs(),"features":requested_features})
        } else {
            json!({"version":protocol,"type":"office.created","pairing_id":id,"verification_code":code,"expires_in":app.inner.config.pairing_ttl.as_secs()})
        },
    );
    Ok(())
}

async fn claim(
    app: &App,
    conn: u64,
    tx: &Tx,
    subject: [u8; 32],
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    let enhanced = protocol == PROTOCOL_V2 && m.contains_key("features");
    let expected: &[&str] = if enhanced {
        &[
            "version",
            "type",
            "verification_code",
            "capabilities",
            "features",
        ]
    } else if protocol == PROTOCOL_V2 {
        &["version", "type", "verification_code", "capabilities"]
    } else {
        &["version", "type", "verification_code"]
    };
    if !exact(&m, expected) {
        return Err("invalid_frame");
    }
    let offered = if protocol == PROTOCOL_V2 {
        capabilities(&m)?
    } else {
        vec!["agent.v1".to_owned()]
    };
    let offered_features = if enhanced {
        features(&m, app.inner.config.pairing_resume_enabled)?
    } else {
        Vec::new()
    };
    let code = string(&m, "verification_code")?;
    if code.len() != 6 || !code.bytes().all(|b| b.is_ascii_digit()) {
        return Err("invalid_code");
    }
    let mut s = app.inner.state.lock().await;
    expire(&app.inner, &mut s);
    let negotiated_claim = s
        .codes
        .get(code)
        .and_then(|id| s.pairings.get(id))
        .is_some_and(|pairing| {
            pairing.version == PROTOCOL_V2
                && pairing.negotiated_subject == Some(subject)
                && pairing.negotiated_capabilities
                    == pairing
                        .requested_capabilities
                        .iter()
                        .filter(|name| offered.contains(name))
                        .cloned()
                        .collect::<Vec<_>>()
                && pairing.negotiated_features
                    == pairing
                        .features
                        .iter()
                        .filter(|name| offered_features.contains(name))
                        .cloned()
                        .collect::<Vec<_>>()
        });
    if !negotiated_claim {
        if s.global_claims
            .0
            .is_none_or(|start| start.elapsed() > app.inner.config.pairing_ttl)
        {
            s.global_claims = (Some(Instant::now()), 0);
        }
        s.global_claims.1 = s.global_claims.1.saturating_add(1);
        if s.global_claims.1 > app.inner.config.max_global_claims {
            return Err("claim_rate_limited");
        }
        if s.claim_attempts.len() >= 10_000 && !s.claim_attempts.contains_key(&subject) {
            return Err("relay_busy");
        }
        let attempts = s
            .claim_attempts
            .entry(subject)
            .or_insert((Instant::now(), 0));
        if attempts.0.elapsed() > app.inner.config.pairing_ttl {
            *attempts = (Instant::now(), 0);
        }
        attempts.1 = attempts.1.saturating_add(1);
        if attempts.1 > app.inner.config.max_claim_attempts {
            return Err("claim_limit");
        }
    }
    let Some(id) = s.codes.get(code).cloned() else {
        return Err("invalid_code");
    };
    let max = app.inner.config.max_claim_attempts;
    let p = s.pairings.get_mut(&id).ok_or("invalid_code")?;
    if p.version != protocol {
        return Err("invalid_frame");
    }
    if p.pc.is_some() {
        return Err("already_claimed");
    }
    let negotiated_capabilities: Vec<_> = p
        .requested_capabilities
        .iter()
        .filter(|name| offered.contains(name))
        .cloned()
        .collect();
    let negotiated_features: Vec<_> = p
        .features
        .iter()
        .filter(|name| offered_features.contains(name))
        .cloned()
        .collect();
    if negotiated_capabilities.is_empty() {
        return Err("capability_not_negotiated");
    }
    if !negotiated_claim {
        p.attempts = p.attempts.saturating_add(1);
        if p.attempts > max {
            return Err("claim_limit");
        }
    }
    p.pc = Some((conn, tx.clone()));
    p.negotiated_capabilities = negotiated_capabilities;
    p.negotiated_features = negotiated_features;
    p.pc_subject = Some(subject);
    p.negotiated_subject = None;
    send(
        tx,
        if enhanced {
            json!({"version":protocol,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs(),"capabilities":p.negotiated_capabilities,"features":p.negotiated_features})
        } else if protocol == PROTOCOL_V2 {
            json!({"version":protocol,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs(),"capabilities":p.negotiated_capabilities})
        } else {
            json!({"version":1,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs()})
        },
    );
    Ok(())
}

async fn approve(
    app: &App,
    conn: u64,
    _tx: &Tx,
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    let enhanced = protocol == PROTOCOL_V2 && m.contains_key("features");
    let expected: &[&str] = if enhanced {
        &["version", "type", "pairing_id", "capabilities", "features"]
    } else if protocol == PROTOCOL_V2 {
        &["version", "type", "pairing_id", "capabilities"]
    } else {
        &["version", "type", "pairing_id"]
    };
    if !exact(&m, expected) {
        return Err("invalid_frame");
    }
    let id = string(&m, "pairing_id")?.to_owned();
    let approved_capabilities = if protocol == PROTOCOL_V2 {
        capabilities(&m)?
    } else {
        vec!["agent.v1".to_owned()]
    };
    let approved_features = if enhanced {
        features(&m, app.inner.config.pairing_resume_enabled)?
    } else {
        Vec::new()
    };
    let mut s = app.inner.state.lock().await;
    expire(&app.inner, &mut s);
    if s.sessions.len() >= 10_000 {
        return Err("relay_busy");
    }
    let pending = s.pairings.get(&id).ok_or("invalid_pairing")?;
    if pending.version != protocol || pending.pc.as_ref().map(|x| x.0) != Some(conn) {
        return Err("invalid_pairing");
    }
    if approved_capabilities != pending.negotiated_capabilities {
        return Err("capability_not_negotiated");
    }
    if approved_features != pending.negotiated_features
        || (!pending.negotiated_features.is_empty() && !enhanced)
    {
        return Err("invalid_frame");
    }
    if pending.negotiated_features == [PAIRING_RESUME] {
        if pending.pending_binding.is_some() {
            return Err("invalid_pairing");
        }
        let binding_id = token();
        let office_tx = pending.office_tx.clone();
        let approved_pc = pending.pc.clone().ok_or("peer_unavailable")?;
        let approved_subject = pending.pc_subject.ok_or("auth_required")?;
        try_send(
            &office_tx,
            json!({
                "version": PROTOCOL_V2,
                "type": "office.binding_offer",
                "pairing_id": id,
                "binding_id": binding_id,
                "capabilities": approved_capabilities,
                "features": [PAIRING_RESUME],
            }),
        )
        .map_err(|_| "peer_unavailable")?;
        let pairing = s.pairings.get_mut(&id).ok_or("invalid_pairing")?;
        pairing.pending_binding = Some(PendingBinding {
            id: binding_id,
            pc: approved_pc.0,
            pc_tx: approved_pc.1,
            subject: approved_subject,
            capabilities: approved_capabilities,
            phase: PendingBindingPhase::Offered,
        });
        let code = pairing.code.clone();
        s.codes.remove(&code);
        return Ok(());
    }
    let p = s.pairings.remove(&id).ok_or("invalid_pairing")?;
    establish_pairing_session(&app.inner, &mut s, p, None, false, None, None)
}

async fn office_binding_result(
    app: &App,
    conn: u64,
    m: Map<String, Value>,
    result: BindingResult,
) -> Result<(), &'static str> {
    if version(&m)? != PROTOCOL_V2 || !exact(&m, &["version", "type", "pairing_id", "binding_id"]) {
        return Err("invalid_frame");
    }
    let pairing_id = string(&m, "pairing_id")?.to_owned();
    let binding_id = string(&m, "binding_id")?.to_owned();
    if !valid_token(&pairing_id) || !valid_token(&binding_id) {
        return Err("invalid_pairing");
    }
    let mut store = app.inner.state.lock().await;
    if store
        .completed_binding_offers
        .get(&pairing_id)
        .is_some_and(|completed| {
            completed.office == conn
                && completed.binding_id == binding_id
                && completed.expires > Instant::now()
        })
    {
        return Ok(());
    }
    let pending = store.pairings.get(&pairing_id).ok_or("invalid_pairing")?;
    let approved = pending.pending_binding.as_ref().ok_or("invalid_pairing")?;
    if pending.office != conn
        || pending.version != PROTOCOL_V2
        || approved.id != binding_id
        || pending.pc.as_ref().map(|pc| pc.0) != Some(approved.pc)
        || pending.pc_subject != Some(approved.subject)
        || pending.negotiated_capabilities != approved.capabilities
        || pending.negotiated_features != [PAIRING_RESUME]
    {
        return Err("invalid_pairing");
    }
    let approved = approved.clone();
    let pending_expires = pending.expires;
    let office_tx = pending.office_tx.clone();
    let public_key = pending.binding_public_key.ok_or("invalid_frame")?;
    let host = pending.host.clone();
    if pending_expires <= Instant::now() {
        let pairing = store
            .pairings
            .remove(&pairing_id)
            .ok_or("invalid_pairing")?;
        let committed = approved.phase == PendingBindingPhase::CommitSent;
        return fallback_staged_pairing(&app.inner, &mut store, pairing, binding_id, committed);
    }
    match result {
        BindingResult::Abort => {
            let pairing = store
                .pairings
                .remove(&pairing_id)
                .ok_or("invalid_pairing")?;
            fallback_staged_pairing(
                &app.inner,
                &mut store,
                pairing,
                binding_id,
                approved.phase == PendingBindingPhase::CommitSent,
            )
        }
        BindingResult::Ready if approved.phase == PendingBindingPhase::CommitSent => {
            if try_send(
                &office_tx,
                json!({"version":PROTOCOL_V2,"type":"office.binding_commit","pairing_id":pairing_id,"binding_id":binding_id}),
            )
            .is_err()
            {
                let pairing = store
                    .pairings
                    .remove(&pairing_id)
                    .ok_or("invalid_pairing")?;
                return fallback_staged_pairing(
                    &app.inner,
                    &mut store,
                    pairing,
                    binding_id,
                    true,
                );
            }
            Ok(())
        }
        BindingResult::Ready => {
            let binding = Binding {
                id: binding_id.clone(),
                subject: approved.subject,
                public_key,
                host: host.clone(),
                origin: OFFICE_ORIGIN.to_owned(),
                capabilities: approved.capabilities.clone(),
            };
            if app.inner.bindings.enroll_pending(&binding).is_err() {
                let pairing = store
                    .pairings
                    .remove(&pairing_id)
                    .ok_or("invalid_pairing")?;
                return fallback_staged_pairing(&app.inner, &mut store, pairing, binding_id, false);
            }
            let pairing = store
                .pairings
                .get_mut(&pairing_id)
                .ok_or("invalid_pairing")?;
            pairing
                .pending_binding
                .as_mut()
                .ok_or("invalid_pairing")?
                .phase = PendingBindingPhase::CommitSent;
            if try_send(
                &pairing.office_tx,
                json!({"version":PROTOCOL_V2,"type":"office.binding_commit","pairing_id":pairing_id,"binding_id":binding_id}),
            )
            .is_err()
            {
                let pairing = store
                    .pairings
                    .remove(&pairing_id)
                    .ok_or("invalid_pairing")?;
                return fallback_staged_pairing(
                    &app.inner,
                    &mut store,
                    pairing,
                    binding_id,
                    true,
                );
            }
            Ok(())
        }
        BindingResult::Committed if approved.phase == PendingBindingPhase::CommitSent => {
            let pairing = store
                .pairings
                .remove(&pairing_id)
                .ok_or("invalid_pairing")?;
            remember_completed_offer(&app.inner, &mut store, pairing_id, conn, binding_id.clone());
            establish_pairing_session(
                &app.inner,
                &mut store,
                pairing,
                Some(binding_id),
                true,
                None,
                Some(approved.subject),
            )
        }
        BindingResult::Committed => Err("invalid_pairing"),
    }
}

fn fallback_staged_pairing(
    inner: &Inner,
    store: &mut Store,
    pairing: Pairing,
    binding_id: String,
    committed: bool,
) -> Result<(), &'static str> {
    remember_completed_offer(
        inner,
        store,
        pairing.id.clone(),
        pairing.office,
        binding_id.clone(),
    );
    if committed {
        compensate_binding(inner, store, &pairing, Some(&binding_id));
    }
    let abort_frame = json!({
        "version": PROTOCOL_V2,
        "type": "office.binding_aborted",
        "pairing_id": pairing.id,
        "binding_id": binding_id,
    });
    establish_pairing_session(inner, store, pairing, None, true, Some(abort_frame), None)
}

fn establish_pairing_session(
    inner: &Inner,
    store: &mut Store,
    pairing: Pairing,
    mut remembered_binding: Option<String>,
    mut explicit_features: bool,
    mut abort_frame: Option<Value>,
    activate_subject: Option<[u8; 32]>,
) -> Result<(), &'static str> {
    store.codes.remove(&pairing.code);
    if store.sessions.len() >= 10_000 {
        compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
        abort_committed_binding(&pairing, remembered_binding.as_deref());
        send_pending_abort(&pairing, abort_frame.as_ref());
        terminate_pairing(&pairing);
        return Err("relay_busy");
    }
    let approved_pc = if let Some(approved) = pairing.pending_binding.as_ref() {
        if pairing.pc.as_ref().map(|pc| pc.0) != Some(approved.pc)
            || pairing.pc_subject != Some(approved.subject)
            || pairing.negotiated_capabilities != approved.capabilities
        {
            compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
            abort_committed_binding(&pairing, remembered_binding.as_deref());
            send_pending_abort(&pairing, abort_frame.as_ref());
            terminate_pairing(&pairing);
            return Err("invalid_pairing");
        }
        Some((approved.pc, approved.pc_tx.clone()))
    } else {
        pairing.pc.clone()
    };
    let Some((pc_conn, pc_tx)) = approved_pc else {
        compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
        abort_committed_binding(&pairing, remembered_binding.as_deref());
        send_pending_abort(&pairing, abort_frame.as_ref());
        terminate_pairing(&pairing);
        return Err("peer_unavailable");
    };
    let pc_sender = pc_tx.sender.clone();
    let office_sender = pairing.office_tx.sender.clone();
    let pc_permit = match pc_sender.try_reserve() {
        Ok(permit) => permit,
        Err(_) => {
            compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
            abort_committed_binding(&pairing, remembered_binding.as_deref());
            send_pending_abort(&pairing, abort_frame.as_ref());
            terminate_pairing(&pairing);
            return Err("peer_unavailable");
        }
    };
    let office_message_count = if abort_frame.is_some() || activate_subject.is_some() {
        2
    } else {
        1
    };
    let mut office_permits = match office_sender.try_reserve_many(office_message_count) {
        Ok(permits) => permits,
        Err(_) => {
            compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
            abort_committed_binding(&pairing, remembered_binding.as_deref());
            send_pending_abort(&pairing, abort_frame.as_ref());
            terminate_pairing(&pairing);
            return Err("peer_unavailable");
        }
    };
    let sid = token();
    let oc = token();
    let pc = token();
    let expires = inner.config.session_ttl.min(inner.config.session_max_ttl);
    let now = Instant::now();
    #[cfg(test)]
    let fail_delivery = inner.fail_approval_delivery.swap(false, Ordering::SeqCst);
    #[cfg(not(test))]
    let fail_delivery = false;
    if fail_delivery {
        drop(pc_permit);
        drop(office_permits);
        compensate_binding(inner, store, &pairing, remembered_binding.as_deref());
        abort_committed_binding(&pairing, remembered_binding.as_deref());
        send_pending_abort(&pairing, abort_frame.as_ref());
        terminate_pairing(&pairing);
        return Err("peer_unavailable");
    }
    if let (Some(binding_id), Some(subject)) = (remembered_binding.as_deref(), activate_subject)
        && !matches!(
            inner.bindings.activate_pending(binding_id, &subject),
            Ok(true)
        )
    {
        let binding_id = binding_id.to_owned();
        compensate_binding(inner, store, &pairing, Some(&binding_id));
        abort_frame = Some(json!({
            "version": PROTOCOL_V2,
            "type": "office.binding_aborted",
            "pairing_id": pairing.id,
            "binding_id": binding_id,
        }));
        remembered_binding = None;
        explicit_features = true;
    }
    let pc_approved = if let Some(binding_id) = remembered_binding.as_ref() {
        json!({"version":pairing.version,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities,"features":[PAIRING_RESUME],"binding_id":binding_id})
    } else if pairing.version == PROTOCOL_V2 && explicit_features {
        json!({"version":pairing.version,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities,"features":[]})
    } else if pairing.version == PROTOCOL_V2 {
        json!({"version":pairing.version,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities})
    } else {
        json!({"version":1,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs()})
    };
    let office_approved = if let Some(binding_id) = remembered_binding.as_ref() {
        json!({"version":pairing.version,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities,"features":[PAIRING_RESUME],"binding_id":binding_id})
    } else if pairing.version == PROTOCOL_V2 && explicit_features {
        json!({"version":pairing.version,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities,"features":[]})
    } else if pairing.version == PROTOCOL_V2 {
        json!({"version":pairing.version,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs(),"capabilities":pairing.negotiated_capabilities})
    } else {
        json!({"version":1,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs()})
    };
    if let Some(abort_frame) = abort_frame {
        office_permits
            .next()
            .expect("reserved binding abort permit")
            .send(Message::Text(abort_frame.to_string().into()));
    }
    pc_permit.send(Message::Text(pc_approved.to_string().into()));
    office_permits
        .next()
        .expect("reserved Office approval permit")
        .send(Message::Text(office_approved.to_string().into()));
    store
        .connection_sessions
        .entry(pairing.office)
        .or_default()
        .insert(sid.clone());
    store
        .connection_sessions
        .entry(pc_conn)
        .or_default()
        .insert(sid.clone());
    store.sessions.insert(
        sid,
        Session {
            version: pairing.version,
            host: pairing.host,
            office: pairing.office,
            office_tx: pairing.office_tx,
            pc: pc_conn,
            pc_tx,
            office_cap: oc,
            pc_cap: pc,
            expires: now + expires,
            absolute_expires: now + inner.config.session_max_ttl,
            active: None,
            used_requests: VecDeque::new(),
            capabilities: pairing.negotiated_capabilities,
            diagnostics: 0,
            diagnostic_window_started: now,
            diagnostic_window_count: 0,
            binding_id: remembered_binding,
        },
    );
    Ok(())
}

fn compensate_binding(
    inner: &Inner,
    store: &mut Store,
    pairing: &Pairing,
    binding_id: Option<&str>,
) {
    let subject = pairing
        .pending_binding
        .as_ref()
        .map(|pending| pending.subject)
        .or(pairing.pc_subject);
    if let (Some(binding_id), Some(subject)) = (binding_id, subject)
        && !matches!(inner.bindings.revoke(binding_id, &subject), Ok(true))
    {
        store.denied_bindings.insert(binding_id.to_owned());
        store
            .pending_revocations
            .insert(binding_id.to_owned(), subject);
    }
}

fn abort_committed_binding(pairing: &Pairing, binding_id: Option<&str>) {
    if let Some(binding_id) = binding_id {
        send(
            &pairing.office_tx,
            json!({
                "version": PROTOCOL_V2,
                "type": "office.binding_aborted",
                "pairing_id": pairing.id,
                "binding_id": binding_id,
            }),
        );
    }
}

fn send_pending_abort(pairing: &Pairing, abort_frame: Option<&Value>) {
    if let Some(abort_frame) = abort_frame {
        send(&pairing.office_tx, abort_frame.clone());
    }
}

fn terminate_pairing(pairing: &Pairing) {
    pairing.office_tx.failed.notify_one();
    if let Some(approved) = pairing.pending_binding.as_ref() {
        approved.pc_tx.failed.notify_one();
    } else if let Some((_, pc_tx)) = pairing.pc.as_ref() {
        pc_tx.failed.notify_one();
    }
}

fn remember_completed_offer(
    inner: &Inner,
    store: &mut Store,
    pairing_id: String,
    office: u64,
    binding_id: String,
) {
    if store.completed_binding_offers.len() >= 10_000 {
        store
            .completed_binding_offers
            .retain(|_, completed| completed.expires > Instant::now());
    }
    if store.completed_binding_offers.len() >= 10_000
        && let Some(oldest) = store
            .completed_binding_offers
            .iter()
            .min_by_key(|(_, completed)| completed.expires)
            .map(|(pairing_id, _)| pairing_id.clone())
    {
        store.completed_binding_offers.remove(&oldest);
    }
    store.completed_binding_offers.insert(
        pairing_id,
        CompletedBindingOffer {
            office,
            binding_id,
            expires: Instant::now() + inner.config.pairing_ttl,
        },
    );
}

async fn office_resume(
    app: &App,
    conn: u64,
    tx: &Tx,
    origin: Option<&str>,
    peer: IpAddr,
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    if !app.inner.config.pairing_resume_enabled
        || version(&m)? != PROTOCOL_V2
        || !exact(
            &m,
            &["version", "type", "binding_id", "host", "capabilities"],
        )
        || origin != Some(OFFICE_ORIGIN)
    {
        return Err("invalid_frame");
    }
    {
        let mut store = app.inner.state.lock().await;
        consume_office_resume_attempt(
            &mut store,
            conn,
            peer,
            app.inner.config.pairing_ttl,
            app.inner.config.max_global_resume_attempts,
        )?;
    }
    let binding_id = string(&m, "binding_id")?;
    if !valid_token(binding_id) {
        return Err("binding_unavailable");
    }
    if app
        .inner
        .state
        .lock()
        .await
        .denied_bindings
        .contains(binding_id)
    {
        return Err("binding_unavailable");
    }
    let host = string(&m, "host")?;
    let requested_capabilities = resume_capabilities(&m)?;
    let binding = app
        .inner
        .bindings
        .get_live(binding_id)
        .map_err(|_| "binding_unavailable")?
        .ok_or("binding_unavailable")?;
    if binding.host != host
        || binding.origin != OFFICE_ORIGIN
        || binding.capabilities != requested_capabilities
    {
        return Err("binding_unavailable");
    }
    let mut challenge_bytes = [0_u8; 32];
    rand::rng().fill(&mut challenge_bytes);
    let challenge = URL_SAFE_NO_PAD.encode(challenge_bytes);
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    if store.resume_challenges.contains_key(&conn) || store.office_resumes.contains_key(&conn) {
        return Err("resume_limit");
    }
    if store.resume_challenges.len() + store.office_resumes.len() >= 10_000 {
        return Err("resume_limit");
    }
    store.resume_challenges.insert(
        conn,
        ResumeChallenge {
            binding,
            challenge: challenge.clone(),
            expires: Instant::now()
                + app
                    .inner
                    .config
                    .resume_challenge_ttl
                    .min(RESUME_CHALLENGE_MAX),
        },
    );
    store.expired_resume_challenges.remove(&conn);
    send(
        tx,
        json!({"version":PROTOCOL_V2,"type":"office.challenge","binding_id":binding_id,"challenge":challenge,"expires_in":app.inner.config.resume_challenge_ttl.min(RESUME_CHALLENGE_MAX).as_secs()}),
    );
    Ok(())
}

async fn office_prove(
    app: &App,
    conn: u64,
    tx: &Tx,
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    if version(&m)? != PROTOCOL_V2
        || !exact(
            &m,
            &["version", "type", "binding_id", "challenge", "signature"],
        )
    {
        return Err("invalid_frame");
    }
    let binding_id = string(&m, "binding_id")?;
    let challenge_value = string(&m, "challenge")?;
    let signature_value = string(&m, "signature")?;
    let mut store = app.inner.state.lock().await;
    let challenge = match store.resume_challenges.remove(&conn) {
        Some(challenge) => challenge,
        None => {
            let matching_expired =
                store
                    .expired_resume_challenges
                    .get(&conn)
                    .is_some_and(|expired| {
                        expired.binding_hash == sha256_bytes(binding_id.as_bytes())
                            && expired.challenge_hash == sha256_bytes(challenge_value.as_bytes())
                            && expired.expires > Instant::now()
                    });
            if matching_expired {
                store.expired_resume_challenges.remove(&conn);
                return Err("challenge_expired");
            }
            return Err("invalid_proof");
        }
    };
    if challenge.expires <= Instant::now() {
        return Err("challenge_expired");
    }
    if challenge.binding.id != binding_id || challenge.challenge != challenge_value {
        return Err("invalid_proof");
    }
    let signature_bytes = URL_SAFE_NO_PAD
        .decode(signature_value)
        .map_err(|_| "invalid_proof")?;
    let signature = Signature::from_slice(&signature_bytes).map_err(|_| "invalid_proof")?;
    let verifying_key = VerifyingKey::from_sec1_bytes(&challenge.binding.public_key)
        .map_err(|_| "invalid_proof")?;
    let transcript = resume_transcript(binding_id, challenge_value, &challenge.binding.host);
    verifying_key
        .verify(transcript.as_bytes(), &signature)
        .map_err(|_| "invalid_proof")?;
    store.office_resumes.insert(
        conn,
        OfficeResume {
            binding: challenge.binding,
            office: conn,
            office_tx: tx.clone(),
            expires: Instant::now() + app.inner.config.pairing_ttl,
        },
    );
    if !complete_resume(app, &mut store, binding_id)? {
        send(
            tx,
            json!({"version":PROTOCOL_V2,"type":"office.waiting_for_pc"}),
        );
    }
    Ok(())
}

async fn pc_resume(
    app: &App,
    conn: u64,
    tx: &Tx,
    subject: [u8; 32],
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    if !app.inner.config.pairing_resume_enabled
        || version(&m)? != PROTOCOL_V2
        || !exact(&m, &["version", "type", "binding_id", "capabilities"])
    {
        return Err("invalid_frame");
    }
    let binding_id = string(&m, "binding_id")?;
    if !valid_token(binding_id) {
        return Err("binding_unavailable");
    }
    if app
        .inner
        .state
        .lock()
        .await
        .denied_bindings
        .contains(binding_id)
    {
        return Err("binding_unavailable");
    }
    let offered_capabilities = resume_capabilities(&m)?;
    let binding = app
        .inner
        .bindings
        .get_live(binding_id)
        .map_err(|_| "binding_unavailable")?
        .ok_or("binding_unavailable")?;
    if binding.subject != subject {
        return Err("binding_unavailable");
    }
    if binding.capabilities != offered_capabilities {
        return Err("capability_not_negotiated");
    }
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    let attempts = store.connection_resume_attempts.entry(conn).or_default();
    *attempts = attempts.saturating_add(1);
    if *attempts > 5 {
        return Err("resume_rate_limited");
    }
    if store.pc_resumes.contains_key(&conn) {
        return Err("resume_limit");
    }
    if store.pc_resumes.len() >= 10_000 {
        return Err("resume_limit");
    }
    store.pc_resumes.insert(
        conn,
        PcResume {
            binding_id: binding.id.clone(),
            pc: conn,
            pc_tx: tx.clone(),
            subject,
            capabilities: offered_capabilities,
            expires: Instant::now() + app.inner.config.pairing_ttl,
        },
    );
    if !complete_resume(app, &mut store, binding_id)? {
        send(
            tx,
            json!({"version":PROTOCOL_V2,"type":"pc.waiting_for_office"}),
        );
    }
    Ok(())
}

fn consume_office_resume_attempt(
    store: &mut Store,
    conn: u64,
    peer: IpAddr,
    window: Duration,
    max_global_attempts: u32,
) -> Result<(), &'static str> {
    if store
        .global_resume_attempts
        .0
        .is_none_or(|start| start.elapsed() > window)
    {
        store.global_resume_attempts = (Some(Instant::now()), 0);
        store
            .resume_attempts
            .retain(|_, (started, _)| started.elapsed() <= window);
    }
    store.global_resume_attempts.1 = store.global_resume_attempts.1.saturating_add(1);
    if store.global_resume_attempts.1 > max_global_attempts {
        return Err("resume_rate_limited");
    }
    if store.resume_attempts.len() >= MAX_TRACKED_RESUME_IPS
        && !store.resume_attempts.contains_key(&peer)
    {
        return Err("relay_busy");
    }
    let ip_attempts = store
        .resume_attempts
        .entry(peer)
        .or_insert((Instant::now(), 0));
    if ip_attempts.0.elapsed() > window {
        *ip_attempts = (Instant::now(), 0);
    }
    ip_attempts.1 = ip_attempts.1.saturating_add(1);
    if ip_attempts.1 > 20 {
        return Err("resume_rate_limited");
    }
    let attempts = store.connection_resume_attempts.entry(conn).or_default();
    *attempts = attempts.saturating_add(1);
    if *attempts > 5 {
        return Err("resume_rate_limited");
    }
    Ok(())
}

fn complete_resume(app: &App, store: &mut Store, binding_id: &str) -> Result<bool, &'static str> {
    if store.denied_bindings.contains(binding_id) {
        return Err("binding_unavailable");
    }
    let office_conn = store
        .office_resumes
        .iter()
        .find(|(_, resume)| resume.binding.id == binding_id)
        .map(|(conn, _)| *conn);
    let pc_conn = store
        .pc_resumes
        .iter()
        .find(|(_, resume)| resume.binding_id == binding_id)
        .map(|(conn, _)| *conn);
    let (Some(office_conn), Some(pc_conn)) = (office_conn, pc_conn) else {
        return Ok(false);
    };
    if store.sessions.len() >= 10_000 {
        return Err("relay_busy");
    }
    let pending_office = store
        .office_resumes
        .get(&office_conn)
        .ok_or("peer_unavailable")?;
    let pending_pc = store.pc_resumes.get(&pc_conn).ok_or("peer_unavailable")?;
    let live_binding = app.inner.bindings.get_live(binding_id).ok().flatten();
    let binding_is_live = live_binding.is_some_and(|binding| {
        binding.subject == pending_pc.subject
            && binding.host == pending_office.binding.host
            && binding.origin == pending_office.binding.origin
            && binding.capabilities == pending_pc.capabilities
    }) && matches!(app.inner.bindings.touch(binding_id), Ok(true));
    if !binding_is_live {
        if let Some(office) = store.office_resumes.remove(&office_conn) {
            versioned_error(&office.office_tx, PROTOCOL_V2, "binding_unavailable");
        }
        if let Some(pc) = store.pc_resumes.remove(&pc_conn) {
            versioned_error(&pc.pc_tx, PROTOCOL_V2, "binding_unavailable");
        }
        return Ok(true);
    }
    let office = store
        .office_resumes
        .remove(&office_conn)
        .ok_or("peer_unavailable")?;
    let pc = store
        .pc_resumes
        .remove(&pc_conn)
        .ok_or("peer_unavailable")?;
    if office.binding.subject != pc.subject || office.binding.capabilities != pc.capabilities {
        return Err("binding_unavailable");
    }
    let session_id = token();
    let office_capability = token();
    let pc_capability = token();
    let expires = app
        .inner
        .config
        .session_ttl
        .min(app.inner.config.session_max_ttl);
    let office_frame = json!({"version":PROTOCOL_V2,"type":"office.approved","session_id":session_id,"capability":office_capability,"expires_in":expires.as_secs(),"capabilities":office.binding.capabilities});
    let pc_frame = json!({"version":PROTOCOL_V2,"type":"pc.approved","session_id":session_id,"capability":pc_capability,"expires_in":expires.as_secs(),"capabilities":office.binding.capabilities});
    if office.office_tx.sender.capacity() == 0
        || pc.pc_tx.sender.capacity() == 0
        || try_send(&office.office_tx, office_frame).is_err()
        || try_send(&pc.pc_tx, pc_frame).is_err()
    {
        office.office_tx.failed.notify_one();
        pc.pc_tx.failed.notify_one();
        return Err("peer_unavailable");
    }
    let now = Instant::now();
    store
        .connection_sessions
        .entry(office.office)
        .or_default()
        .insert(session_id.clone());
    store
        .connection_sessions
        .entry(pc.pc)
        .or_default()
        .insert(session_id.clone());
    store.sessions.insert(
        session_id,
        Session {
            version: PROTOCOL_V2,
            host: office.binding.host,
            office: office.office,
            office_tx: office.office_tx,
            pc: pc.pc,
            pc_tx: pc.pc_tx,
            office_cap: office_capability,
            pc_cap: pc_capability,
            expires: now + expires,
            absolute_expires: now + app.inner.config.session_max_ttl,
            active: None,
            used_requests: VecDeque::new(),
            capabilities: office.binding.capabilities,
            diagnostics: 0,
            diagnostic_window_started: now,
            diagnostic_window_count: 0,
            binding_id: Some(binding_id.to_owned()),
        },
    );
    Ok(true)
}

async fn revoke_binding(
    app: &App,
    _conn: u64,
    tx: &Tx,
    subject: [u8; 32],
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    if version(&m)? != PROTOCOL_V2 || !exact(&m, &["version", "type", "binding_id"]) {
        return Err("invalid_frame");
    }
    let binding_id = string(&m, "binding_id")?;
    if !valid_token(binding_id)
        || !app
            .inner
            .bindings
            .revoke(binding_id, &subject)
            .map_err(|_| "binding_unavailable")?
    {
        return Err("binding_unavailable");
    }
    let mut store = app.inner.state.lock().await;
    store.denied_bindings.remove(binding_id);
    store.pending_revocations.remove(binding_id);
    store
        .resume_challenges
        .retain(|_, challenge| challenge.binding.id != binding_id);
    let waiting_office: Vec<u64> = store
        .office_resumes
        .iter()
        .filter(|(_, resume)| resume.binding.id == binding_id)
        .map(|(conn, _)| *conn)
        .collect();
    for conn in waiting_office {
        if let Some(resume) = store.office_resumes.remove(&conn) {
            versioned_error(&resume.office_tx, PROTOCOL_V2, "binding_unavailable");
        }
    }
    let waiting_pc: Vec<u64> = store
        .pc_resumes
        .iter()
        .filter(|(_, resume)| resume.binding_id == binding_id)
        .map(|(conn, _)| *conn)
        .collect();
    for conn in waiting_pc {
        if let Some(resume) = store.pc_resumes.remove(&conn) {
            versioned_error(&resume.pc_tx, PROTOCOL_V2, "binding_unavailable");
        }
    }
    send(
        tx,
        json!({"version":PROTOCOL_V2,"type":"pc.binding_revoked","binding_id":binding_id}),
    );
    let sessions: Vec<String> = store
        .sessions
        .iter()
        .filter(|(_, session)| session.binding_id.as_deref() == Some(binding_id))
        .map(|(id, _)| id.clone())
        .collect();
    for session_id in sessions {
        if let Some(session) = store.sessions.remove(&session_id) {
            if let Some(ids) = store.connection_sessions.get_mut(&session.office) {
                ids.remove(&session_id);
            }
            if let Some(ids) = store.connection_sessions.get_mut(&session.pc) {
                ids.remove(&session_id);
            }
            if let Some(active) = session.active.as_ref() {
                send(
                    &session.pc_tx,
                    json!({"version":PROTOCOL_V2,"type":"relay.cancel","session_id":session_id,"request_id":active.id}),
                );
            }
            versioned_error(&session.office_tx, PROTOCOL_V2, "session_revoked");
            versioned_error(&session.pc_tx, PROTOCOL_V2, "session_revoked");
        }
    }
    Ok(())
}

fn valid_token(value: &str) -> bool {
    value.len() == 43 && value.bytes().all(|byte| byte.is_ascii_alphanumeric())
}

fn resume_transcript(binding_id: &str, challenge: &str, host: &str) -> String {
    format!("wiswork-office-resume-v1\n{binding_id}\n{challenge}\n{OFFICE_ORIGIN}\n{host}")
}

fn sha256_bytes(value: &[u8]) -> [u8; 32] {
    Sha256::digest(value).into()
}

const DIAGNOSTIC_REQUIRED_KEYS: &[&str] = &[
    "version",
    "type",
    "session_id",
    "capability",
    "event_id",
    "trace_id",
    "timestamp_ms",
    "host",
    "platform",
    "build",
    "tool",
    "phase",
    "outcome",
    "error_code",
    "duration_ms",
    "requirement_sets",
];
const DIAGNOSTIC_OPTIONAL_KEYS: &[&str] = &[
    "office_error_code",
    "office_error_name",
    "office_error_location",
];
const DIAGNOSTIC_ERROR_CODES: &[&str] = &[
    "office_api_unsupported",
    "office_read_failed",
    "office_write_failed",
    "office_verify_failed",
    "office_recovery_failed",
    "proposal_missing",
    "proposal_stale",
    "auth_required",
    "network_error",
    "provider_unavailable",
    "request_timeout",
    "agent_run_failed",
    "cancelled",
    "diagnostic_upload_failed",
    "user_rejected_change",
];

fn diagnostic_keys_are_exact(m: &Map<String, Value>) -> bool {
    DIAGNOSTIC_REQUIRED_KEYS
        .iter()
        .all(|key| m.contains_key(*key))
        && m.keys().all(|key| {
            DIAGNOSTIC_REQUIRED_KEYS.contains(&key.as_str())
                || DIAGNOSTIC_OPTIONAL_KEYS.contains(&key.as_str())
        })
}

fn uuid(value: &str) -> bool {
    value.len() == 36
        && value.bytes().enumerate().all(|(index, byte)| {
            if matches!(index, 8 | 13 | 18 | 23) {
                byte == b'-'
            } else {
                byte.is_ascii_hexdigit()
            }
        })
}

fn bounded_identifier(value: &str, maximum: usize) -> bool {
    !value.is_empty()
        && value.len() <= maximum
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(
                    byte,
                    b'.' | b'_' | b'-' | b':' | b'/' | b'(' | b')' | b'[' | b']'
                )
        })
}

fn optional_identifier(m: &Map<String, Value>, key: &str) -> Result<Option<String>, &'static str> {
    match m.get(key) {
        None => Ok(None),
        Some(value) => {
            let value = value.as_str().ok_or("invalid_frame")?;
            if !bounded_identifier(value, 128) {
                return Err("invalid_frame");
            }
            Ok(Some(value.to_owned()))
        }
    }
}

async fn diagnostic(
    app: &App,
    conn: u64,
    tx: &Tx,
    m: Map<String, Value>,
    wire_size: usize,
) -> Result<(), &'static str> {
    if wire_size > DIAGNOSTIC_MAX {
        return Err("diagnostic_too_large");
    }
    if version(&m)? != PROTOCOL_V2 || !diagnostic_keys_are_exact(&m) {
        return Err("invalid_frame");
    }
    let sid = string(&m, "session_id")?;
    let cap = string(&m, "capability")?;
    let event_id = string(&m, "event_id")?;
    let trace_id = string(&m, "trace_id")?;
    if !uuid(event_id) || !uuid(trace_id) {
        return Err("invalid_frame");
    }
    let timestamp_ms = m
        .get("timestamp_ms")
        .and_then(Value::as_u64)
        .ok_or("invalid_frame")?;
    let duration_ms = m
        .get("duration_ms")
        .and_then(Value::as_u64)
        .filter(|value| *value <= 86_400_000)
        .ok_or("invalid_frame")?;
    let host = string(&m, "host")?;
    if !matches!(host, "word" | "excel" | "powerpoint") {
        return Err("invalid_frame");
    }
    let platform = string(&m, "platform")?;
    if !matches!(
        platform,
        "pc" | "mac" | "office_online" | "ios" | "android" | "universal" | "unknown"
    ) {
        return Err("invalid_frame");
    }
    let build = string(&m, "build")?;
    let tool = string(&m, "tool")?;
    if !bounded_identifier(build, 96) || !(tool == "unknown" || bounded_identifier(tool, 96)) {
        return Err("invalid_frame");
    }
    let phase = string(&m, "phase")?;
    if !matches!(
        phase,
        "tool" | "proposal" | "validate" | "write" | "verify" | "recovery" | "transport"
    ) {
        return Err("invalid_frame");
    }
    let outcome = string(&m, "outcome")?;
    if !matches!(outcome, "failed" | "unsupported" | "cancelled") {
        return Err("invalid_frame");
    }
    let error_code = string(&m, "error_code")?;
    if !DIAGNOSTIC_ERROR_CODES.contains(&error_code) {
        return Err("invalid_frame");
    }
    let office_error_code = optional_identifier(&m, "office_error_code")?;
    let office_error_name = optional_identifier(&m, "office_error_name")?;
    let office_error_location = optional_identifier(&m, "office_error_location")?;
    let requirement_sets = m
        .get("requirement_sets")
        .and_then(Value::as_object)
        .ok_or("invalid_frame")?;
    if requirement_sets.len() > 4
        || requirement_sets.iter().any(|(name, value)| {
            !matches!(
                name.as_str(),
                "OfficeApi" | "WordApi" | "ExcelApi" | "PowerPointApi"
            ) || !value.is_boolean()
        })
    {
        return Err("invalid_frame");
    }

    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    let session = store.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap || session.version != PROTOCOL_V2 {
        return Err("invalid_capability");
    }
    let expected_host = match session.host.as_str() {
        "Word" => "word",
        "Excel" => "excel",
        "PowerPoint" => "powerpoint",
        _ => return Err("diagnostic_host_mismatch"),
    };
    if host != expected_host {
        return Err("diagnostic_host_mismatch");
    }
    if session.diagnostics >= DIAGNOSTIC_SESSION_MAX {
        return Err("diagnostic_limit");
    }
    if session.diagnostic_window_started.elapsed() >= app.inner.config.diagnostic_window {
        session.diagnostic_window_started = Instant::now();
        session.diagnostic_window_count = 0;
    }
    if session.diagnostic_window_count >= app.inner.config.max_diagnostics_per_window {
        return Err("diagnostic_rate_limited");
    }
    session.diagnostics += 1;
    session.diagnostic_window_count += 1;

    let log_entry = json!({
        "event": "office_diagnostic",
        "connection_id": conn,
        "session_id": sid,
        "event_id": event_id,
        "trace_id": trace_id,
        "timestamp_ms": timestamp_ms,
        "host": host,
        "platform": platform,
        "build": build,
        "tool": tool,
        "phase": phase,
        "outcome": outcome,
        "error_code": error_code,
        "office_error_code": office_error_code,
        "office_error_name": office_error_name,
        "office_error_location": office_error_location,
        "duration_ms": duration_ms,
        "requirement_sets": requirement_sets,
    });
    drop(store);
    eprintln!("{log_entry}");
    diagnostic_response(
        tx,
        json!({
            "version": PROTOCOL_V2,
            "type": "office.diagnostic.accepted",
            "event_id": event_id,
        }),
    );
    Ok(())
}

async fn reject(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(&m, &["version", "type", "pairing_id"]) {
        return Err("invalid_frame");
    }
    let id = string(&m, "pairing_id")?;
    let mut s = app.inner.state.lock().await;
    let p = s.pairings.get(id).ok_or("invalid_pairing")?;
    if p.pc.as_ref().map(|v| v.0) != Some(conn) || p.version != protocol {
        return Err("invalid_pairing");
    }
    let p = s.pairings.remove(id).unwrap();
    s.codes.remove(&p.code);
    if let Some(approved) = p.pending_binding.as_ref() {
        if approved.phase == PendingBindingPhase::CommitSent {
            compensate_binding(&app.inner, &mut s, &p, Some(&approved.id));
        }
        remember_completed_offer(
            &app.inner,
            &mut s,
            p.id.clone(),
            p.office,
            approved.id.clone(),
        );
        send(
            &p.office_tx,
            json!({"version":PROTOCOL_V2,"type":"office.binding_aborted","pairing_id":p.id,"binding_id":approved.id}),
        );
    }
    send(
        &p.office_tx,
        json!({"version":protocol,"type":"office.rejected"}),
    );
    Ok(())
}

fn session_fields(m: &Map<String, Value>) -> Result<(&str, &str, &str), &'static str> {
    Ok((
        string(m, "session_id")?,
        string(m, "capability")?,
        string(m, "request_id")?,
    ))
}
fn valid_identifier(value: &str) -> bool {
    (8..=128).contains(&value.len())
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-'))
}
async fn request(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    let mut expected = vec![
        "version",
        "type",
        "session_id",
        "capability",
        "request_id",
        "body",
    ];
    if protocol == PROTOCOL_V2 {
        expected.push("capability_name");
    }
    if !exact(&m, &expected) {
        return Err("invalid_frame");
    }
    if serde_json::to_vec(&m["body"])
        .map_err(|_| "invalid_frame")?
        .len()
        > REQUEST_MAX
    {
        return Err("request_too_large");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    if rid.is_empty() || rid.len() > 128 {
        return Err("invalid_frame");
    }
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    let capability_name = if protocol == PROTOCOL_V2 {
        let name = string(&m, "capability_name")?;
        if !session.capabilities.iter().any(|item| item == name) {
            send(
                &session.office_tx,
                json!({"version":protocol,"type":"relay.error","session_id":sid,"request_id":rid,"code":"capability_not_negotiated"}),
            );
            return Ok(());
        }
        Some(name.to_owned())
    } else {
        None
    };
    if session.active.is_some() {
        return Err("request_active");
    }
    if session.used_requests.iter().any(|used| used == rid) {
        return Err("duplicate_request");
    }
    renew_session(session, app.inner.config.session_ttl);
    if session.used_requests.len() == 256 {
        session.used_requests.pop_front();
    }
    session.used_requests.push_back(rid.to_owned());
    session.active = Some(Active {
        id: rid.into(),
        sequence: 0,
        bytes: 0,
        deadline: Instant::now() + app.inner.config.request_ttl,
        started: false,
        pending_tool: None,
        used_tool_calls: VecDeque::new(),
    });
    send(
        &session.pc_tx,
        if let Some(name) = capability_name {
            json!({"version":protocol,"type":"relay.request","session_id":sid,"request_id":rid,"capability_name":name,"body":m["body"]})
        } else {
            json!({"version":1,"type":"relay.request","session_id":sid,"request_id":rid,"body":m["body"]})
        },
    );
    let deadline_app = app.clone();
    let deadline_sid = sid.to_owned();
    let deadline_rid = rid.to_owned();
    let deadline = app.inner.config.request_ttl;
    tokio::spawn(async move {
        tokio::time::sleep(deadline).await;
        let mut store = deadline_app.inner.state.lock().await;
        if let Some(session) = store.sessions.get_mut(&deadline_sid)
            && session
                .active
                .as_ref()
                .is_some_and(|a| a.id == deadline_rid)
        {
            session.active = None;
            let protocol = session.version;
            send(
                &session.pc_tx,
                json!({"version":protocol,"type":"relay.cancel","session_id":deadline_sid,"request_id":deadline_rid}),
            );
            send(
                &session.office_tx,
                json!({"version":protocol,"type":"relay.error","session_id":deadline_sid,"request_id":deadline_rid,"code":"request_timeout"}),
            );
        }
    });
    Ok(())
}
async fn cancel(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(
        &m,
        &["version", "type", "session_id", "capability", "request_id"],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|a| a.id.as_str()) != Some(rid) {
        if known_inactive_request(session, rid) {
            return Ok(());
        }
        return Err("invalid_request");
    }
    renew_session(session, app.inner.config.session_ttl);
    session.active = None;
    send(
        &session.pc_tx,
        json!({"version":protocol,"type":"relay.cancel","session_id":sid,"request_id":rid}),
    );
    Ok(())
}

async fn tool_call(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if protocol != PROTOCOL_V2
        || !exact(
            &m,
            &[
                "version",
                "type",
                "session_id",
                "capability",
                "request_id",
                "turn_id",
                "call_id",
                "generation",
                "tool_name",
                "input",
            ],
        )
    {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let turn_id = string(&m, "turn_id")?;
    let call_id = string(&m, "call_id")?;
    let tool_name = string(&m, "tool_name")?;
    let generation = m["generation"].as_u64().ok_or("invalid_frame")?;
    if !valid_identifier(turn_id)
        || !valid_identifier(call_id)
        || tool_name.is_empty()
        || tool_name.len() > 128
        || !tool_name
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-'))
    {
        return Err("invalid_frame");
    }
    if !m["input"].is_object()
        || serde_json::to_vec(&m["input"])
            .map_err(|_| "invalid_frame")?
            .len()
            > REQUEST_MAX
    {
        return Err("request_too_large");
    }
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    let active = session.active.as_mut().ok_or("invalid_request")?;
    if active.id != rid || !active.started || active.deadline <= Instant::now() {
        return Err("invalid_request");
    }
    if active.pending_tool.is_some() {
        return Err("tool_active");
    }
    if active.used_tool_calls.iter().any(|used| used == call_id) {
        return Err("duplicate_tool_call");
    }
    if active.used_tool_calls.len() == 64 {
        active.used_tool_calls.pop_front();
    }
    active.used_tool_calls.push_back(call_id.to_owned());
    active.pending_tool = Some(PendingTool {
        turn_id: turn_id.to_owned(),
        call_id: call_id.to_owned(),
        generation,
    });
    renew_session(session, app.inner.config.session_ttl);
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.tool_call","session_id":sid,"request_id":rid,"turn_id":turn_id,"call_id":call_id,"generation":generation,"tool_name":tool_name,"input":m["input"]}),
    );
    Ok(())
}

async fn tool_result(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if protocol != PROTOCOL_V2
        || !exact(
            &m,
            &[
                "version",
                "type",
                "session_id",
                "capability",
                "request_id",
                "turn_id",
                "call_id",
                "generation",
                "output",
                "is_error",
            ],
        )
    {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let turn_id = string(&m, "turn_id")?;
    let call_id = string(&m, "call_id")?;
    let generation = m["generation"].as_u64().ok_or("invalid_frame")?;
    let output = string(&m, "output")?;
    let is_error = m["is_error"].as_bool().ok_or("invalid_frame")?;
    if output.len() > RESPONSE_MAX {
        return Err("response_too_large");
    }
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    let active = session.active.as_mut().ok_or("invalid_request")?;
    if active.id != rid {
        return Err("invalid_request");
    }
    let pending = active.pending_tool.as_ref().ok_or("invalid_tool_call")?;
    if pending.turn_id != turn_id || pending.call_id != call_id || pending.generation != generation
    {
        return Err("invalid_tool_call");
    }
    active.pending_tool = None;
    renew_session(session, app.inner.config.session_ttl);
    send(
        &session.pc_tx,
        json!({"version":protocol,"type":"relay.tool_result","session_id":sid,"request_id":rid,"turn_id":turn_id,"call_id":call_id,"generation":generation,"output":output,"is_error":is_error}),
    );
    Ok(())
}

async fn chunk(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(
        &m,
        &[
            "version",
            "type",
            "session_id",
            "capability",
            "request_id",
            "sequence",
            "data",
        ],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let seq = m["sequence"].as_u64().ok_or("invalid_frame")?;
    let data = string(&m, "data")?;
    let decoded = STANDARD.decode(data).map_err(|_| "invalid_chunk")?;
    if decoded.len() > CHUNK_MAX {
        return Err("chunk_too_large");
    }
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if known_inactive_request(session, rid) {
        return Ok(());
    }
    let active = session.active.as_ref().ok_or("invalid_request")?;
    if active.id != rid || !active.started || active.sequence != seq {
        return Err("invalid_sequence");
    }
    if active.deadline <= Instant::now() {
        return Err("request_timeout");
    }
    let next_bytes = active
        .bytes
        .checked_add(decoded.len())
        .ok_or("response_too_large")?;
    if next_bytes > RESPONSE_MAX {
        return Err("response_too_large");
    }
    renew_session(session, app.inner.config.session_ttl);
    let active = session.active.as_mut().ok_or("invalid_request")?;
    active.sequence += 1;
    active.bytes = next_bytes;
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.chunk","session_id":sid,"request_id":rid,"sequence":seq,"data":data}),
    );
    Ok(())
}
async fn start(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(
        &m,
        &[
            "version",
            "type",
            "session_id",
            "capability",
            "request_id",
            "status",
            "content_type",
        ],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let status = m["status"]
        .as_u64()
        .filter(|v| (200..=599).contains(v) && !matches!(v, 204 | 205 | 304))
        .ok_or("invalid_status")?;
    let content_type = string(&m, "content_type")?;
    if content_type.is_empty()
        || content_type.len() > 128
        || content_type.bytes().any(|b| !(0x20..=0x7e).contains(&b))
    {
        return Err("invalid_content_type");
    }
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    let session = store.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if known_inactive_request(session, rid) {
        return Ok(());
    }
    let active = session.active.as_ref().ok_or("invalid_request")?;
    if active.id != rid || active.started {
        return Err("invalid_request");
    }
    renew_session(session, app.inner.config.session_ttl);
    let active = session.active.as_mut().ok_or("invalid_request")?;
    active.started = true;
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.start","session_id":sid,"request_id":rid,"status":status,"content_type":content_type}),
    );
    Ok(())
}

fn valid_enhanced_statement(value: &Value, host: &str) -> bool {
    let Some(record) = value.as_object() else {
        return false;
    };
    let expected_host = match host {
        "Word" => "office-word",
        "Excel" => "office-excel",
        "PowerPoint" => "office-powerpoint",
        _ => return false,
    };
    exact(
        record,
        &[
            "version",
            "runtime_mode",
            "runtime_instance",
            "component_version",
            "host",
            "raw_office",
            "expires_at",
            "policy_generation",
            "session_generation",
        ],
    ) && record["version"] == 1
        && record["runtime_mode"] == "enhanced"
        && record["runtime_instance"]
            .as_str()
            .is_some_and(valid_identifier)
        && record["component_version"]
            .as_str()
            .is_some_and(|v| v == "0.147.0")
        && record["host"] == expected_host
        && record["raw_office"].is_boolean()
        && record["expires_at"].as_u64().is_some_and(|v| v > 0)
        && record["policy_generation"].as_u64().is_some()
        && record["session_generation"].as_u64().is_some()
}

async fn session_state(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if protocol != PROTOCOL_V2
        || !exact(
            &m,
            &[
                "version",
                "type",
                "session_id",
                "capability",
                "generation",
                "enhanced",
            ],
        )
    {
        return Err("invalid_frame");
    }
    let sid = string(&m, "session_id")?;
    let cap = string(&m, "capability")?;
    let generation = m["generation"].as_u64().ok_or("invalid_frame")?;
    if !m["enhanced"].is_null()
        && serde_json::to_vec(&m["enhanced"])
            .map_err(|_| "invalid_frame")?
            .len()
            > 4096
    {
        return Err("frame_too_large");
    }
    let mut store = app.inner.state.lock().await;
    expire(&app.inner, &mut store);
    let session = store.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if !m["enhanced"].is_null() && !valid_enhanced_statement(&m["enhanced"], &session.host) {
        return Err("invalid_frame");
    }
    renew_session(session, app.inner.config.session_ttl);
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.session_state","session_id":sid,"generation":generation,"enhanced":m["enhanced"]}),
    );
    Ok(())
}
async fn done(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(
        &m,
        &["version", "type", "session_id", "capability", "request_id"],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|active| active.id.as_str()) != Some(rid) {
        if known_inactive_request(session, rid) {
            return Ok(());
        }
        return Err("invalid_request");
    }
    if !session.active.as_ref().is_some_and(|active| active.started) {
        return Err("invalid_request");
    }
    if session
        .active
        .as_ref()
        .is_some_and(|active| active.pending_tool.is_some())
    {
        return Err("tool_active");
    }
    renew_session(session, app.inner.config.session_ttl);
    session.active = None;
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.done","session_id":sid,"request_id":rid}),
    );
    Ok(())
}
async fn pc_error(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    if !exact(
        &m,
        &[
            "version",
            "type",
            "session_id",
            "capability",
            "request_id",
            "code",
        ],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let code = string(&m, "code")?;
    if !matches!(
        code,
        "upstream_error" | "auth_required" | "quota_exceeded" | "request_failed" | "cancelled"
    ) {
        return Err("invalid_frame");
    }
    let mut st = app.inner.state.lock().await;
    expire(&app.inner, &mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|active| active.id.as_str()) != Some(rid) {
        if known_inactive_request(session, rid) {
            return Ok(());
        }
        return Err("invalid_request");
    }
    renew_session(session, app.inner.config.session_ttl);
    session.active = None;
    send(
        &session.office_tx,
        json!({"version":protocol,"type":"relay.error","session_id":sid,"request_id":rid,"code":code}),
    );
    Ok(())
}

fn expire(inner: &Inner, s: &mut Store) {
    let now = Instant::now();
    let pending_revocations: Vec<_> = s
        .pending_revocations
        .iter()
        .map(|(binding_id, subject)| (binding_id.clone(), *subject))
        .collect();
    for (binding_id, subject) in pending_revocations {
        let converged = matches!(inner.bindings.revoke(&binding_id, &subject), Ok(true))
            || matches!(inner.bindings.get_live(&binding_id), Ok(None));
        if converged {
            s.pending_revocations.remove(&binding_id);
            s.denied_bindings.remove(&binding_id);
        }
    }
    let expired_challenges: Vec<_> = s
        .resume_challenges
        .iter()
        .filter(|(_, challenge)| challenge.expires <= now)
        .map(|(conn, _)| *conn)
        .collect();
    for conn in expired_challenges {
        if let Some(challenge) = s.resume_challenges.remove(&conn) {
            s.expired_resume_challenges.insert(
                conn,
                ExpiredResumeChallenge {
                    binding_hash: sha256_bytes(challenge.binding.id.as_bytes()),
                    challenge_hash: sha256_bytes(challenge.challenge.as_bytes()),
                    expires: now + inner.config.pairing_ttl,
                },
            );
        }
    }
    s.expired_resume_challenges
        .retain(|_, challenge| challenge.expires > now);
    s.completed_binding_offers
        .retain(|_, completed| completed.expires > now);
    let expired_office_resumes: Vec<_> = s
        .office_resumes
        .iter()
        .filter(|(_, resume)| resume.expires <= now)
        .map(|(conn, _)| *conn)
        .collect();
    for conn in expired_office_resumes {
        if let Some(resume) = s.office_resumes.remove(&conn) {
            versioned_error(&resume.office_tx, PROTOCOL_V2, "peer_unavailable");
        }
    }
    let expired_pc_resumes: Vec<_> = s
        .pc_resumes
        .iter()
        .filter(|(_, resume)| resume.expires <= now)
        .map(|(conn, _)| *conn)
        .collect();
    for conn in expired_pc_resumes {
        if let Some(resume) = s.pc_resumes.remove(&conn) {
            versioned_error(&resume.pc_tx, PROTOCOL_V2, "peer_unavailable");
        }
    }
    let dead: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, p)| p.expires <= now)
        .map(|(id, _)| id.clone())
        .collect();
    for id in dead {
        if let Some(p) = s.pairings.remove(&id) {
            if let Some(binding_id) = p.pending_binding.as_ref().map(|pending| pending.id.clone()) {
                let committed = p
                    .pending_binding
                    .as_ref()
                    .is_some_and(|pending| pending.phase == PendingBindingPhase::CommitSent);
                let _ = fallback_staged_pairing(inner, s, p, binding_id, committed);
            } else {
                s.codes.remove(&p.code);
                send(
                    &p.office_tx,
                    json!({"version":p.version,"type":"office.expired"}),
                );
            }
        }
    }
    let expired_sessions: Vec<_> = s
        .sessions
        .iter()
        .filter(|(_, x)| x.expires <= now)
        .map(|(id, _)| id.clone())
        .collect();
    for id in expired_sessions {
        if let Some(x) = s.sessions.remove(&id) {
            if let Some(ids) = s.connection_sessions.get_mut(&x.office) {
                ids.remove(&id);
            }
            if let Some(ids) = s.connection_sessions.get_mut(&x.pc) {
                ids.remove(&id);
            }
            versioned_error(&x.office_tx, x.version, "session_expired");
            versioned_error(&x.pc_tx, x.version, "session_expired");
            if let Some(active) = x.active {
                send(
                    &x.pc_tx,
                    json!({"version":x.version,"type":"relay.cancel","session_id":id,"request_id":active.id}),
                );
            }
        }
    }
    s.claim_attempts
        .retain(|_, (started, _)| started.elapsed() <= inner.config.pairing_ttl);
    s.create_attempts
        .retain(|_, (started, _)| started.elapsed() <= inner.config.pairing_ttl);
    s.preauth_attempts
        .retain(|_, (started, _)| started.elapsed() <= inner.config.pairing_ttl);
    s.resume_attempts
        .retain(|_, (started, _)| started.elapsed() <= inner.config.pairing_ttl);
}
async fn cleanup(app: &App, conn: u64) {
    let mut s = app.inner.state.lock().await;
    s.connection_pairings.remove(&conn);
    s.connection_sessions.remove(&conn);
    s.connection_resume_attempts.remove(&conn);
    s.resume_challenges.remove(&conn);
    s.expired_resume_challenges.remove(&conn);
    s.completed_binding_offers
        .retain(|_, completed| completed.office != conn);
    s.office_resumes.remove(&conn);
    s.pc_resumes.remove(&conn);
    let pids: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, p)| p.office == conn)
        .map(|(id, _)| id.clone())
        .collect();
    for id in pids {
        if let Some(p) = s.pairings.remove(&id) {
            s.codes.remove(&p.code);
            if let Some(approved) = p
                .pending_binding
                .as_ref()
                .filter(|approved| approved.phase == PendingBindingPhase::CommitSent)
            {
                compensate_binding(&app.inner, &mut s, &p, Some(&approved.id));
            }
            if let Some((_, pc_tx)) = p.pc {
                versioned_error(&pc_tx, p.version, "peer_unavailable");
            }
        }
    }
    let staged_pc_pairings: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, pairing)| {
            pairing
                .pending_binding
                .as_ref()
                .is_some_and(|approved| approved.pc == conn)
        })
        .map(|(id, _)| id.clone())
        .collect();
    for id in staged_pc_pairings {
        if let Some(pairing) = s.pairings.remove(&id) {
            s.codes.remove(&pairing.code);
            if let Some(approved) = pairing.pending_binding.as_ref() {
                if approved.phase == PendingBindingPhase::CommitSent {
                    compensate_binding(&app.inner, &mut s, &pairing, Some(&approved.id));
                }
                remember_completed_offer(
                    &app.inner,
                    &mut s,
                    pairing.id.clone(),
                    pairing.office,
                    approved.id.clone(),
                );
                send(
                    &pairing.office_tx,
                    json!({"version":PROTOCOL_V2,"type":"office.binding_aborted","pairing_id":pairing.id,"binding_id":approved.id}),
                );
                send(
                    &pairing.office_tx,
                    json!({"version":PROTOCOL_V2,"type":"office.pc_offline"}),
                );
            }
        }
    }
    for pairing in s.pairings.values_mut() {
        if pairing.pc.as_ref().is_some_and(|value| value.0 == conn) {
            pairing.pc = None;
            send(
                &pairing.office_tx,
                json!({"version":pairing.version,"type":"office.pc_offline"}),
            );
        }
    }
    let ids: Vec<_> = s
        .sessions
        .iter()
        .filter(|(_, x)| x.office == conn || x.pc == conn)
        .map(|(id, _)| id.clone())
        .collect();
    for id in ids {
        if let Some(x) = s.sessions.remove(&id) {
            let peer = if x.office == conn { x.pc } else { x.office };
            if let Some(session_ids) = s.connection_sessions.get_mut(&peer) {
                session_ids.remove(&id);
            }
            if x.office != conn {
                versioned_error(&x.office_tx, x.version, "session_revoked")
            }
            if x.pc != conn {
                send(
                    &x.pc_tx,
                    json!({"version":x.version,"type":"relay.error","code":"session_revoked"}),
                );
            }
            if x.pc != conn
                && let Some(a) = x.active
            {
                send(
                    &x.pc_tx,
                    json!({"version":x.version,"type":"relay.cancel","session_id":id,"request_id":a.id}),
                );
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn test_app() -> App {
        test_app_with_config(Config::default())
    }

    fn test_app_with_config(config: Config) -> App {
        test_app_with_bindings(config, BindingStore::open(None).unwrap())
    }

    fn test_app_with_bindings(config: Config, bindings: BindingStore) -> App {
        App {
            inner: Arc::new(Inner {
                state: Mutex::new(Store::default()),
                next: AtomicU64::new(1),
                config,
                http: reqwest::Client::builder()
                    .no_proxy()
                    .build()
                    .expect("test HTTP client"),
                auth_slots: Arc::new(Semaphore::new(1)),
                bindings,
                fail_approval_delivery: std::sync::atomic::AtomicBool::new(false),
            }),
        }
    }

    async fn insert_test_pairing(app: &App, office_tx: Tx, pc_tx: Tx, subject: [u8; 32]) {
        let signing_key = p256::ecdsa::SigningKey::from_slice(&[21_u8; 32]).unwrap();
        let public_key: [u8; 65] = signing_key
            .verifying_key()
            .to_encoded_point(false)
            .as_bytes()
            .try_into()
            .unwrap();
        let pairing_id = "A".repeat(43);
        let mut store = app.inner.state.lock().await;
        store.codes.insert("123456".to_owned(), pairing_id.clone());
        store.pairings.insert(
            pairing_id.clone(),
            Pairing {
                version: PROTOCOL_V2,
                id: pairing_id,
                code: "123456".to_owned(),
                host: "Word".to_owned(),
                office: 1,
                office_tx,
                pc: Some((2, pc_tx)),
                expires: Instant::now() + Duration::from_secs(60),
                attempts: 0,
                requested_capabilities: vec!["agent.v1".to_owned()],
                negotiated_capabilities: vec!["agent.v1".to_owned()],
                negotiated_subject: None,
                features: vec![PAIRING_RESUME.to_owned()],
                negotiated_features: vec![PAIRING_RESUME.to_owned()],
                binding_public_key: Some(public_key),
                pc_subject: Some(subject),
                pending_binding: None,
            },
        );
    }

    fn test_approval_frame() -> Map<String, Value> {
        json!({
            "version": 2,
            "type": "pc.approve",
            "pairing_id": "A".repeat(43),
            "capabilities": ["agent.v1"],
            "features": [PAIRING_RESUME],
        })
        .as_object()
        .unwrap()
        .clone()
    }

    async fn test_binding_result_frame(app: &App, result: BindingResult) -> Map<String, Value> {
        let binding_id = app
            .inner
            .state
            .lock()
            .await
            .pairings
            .get(&"A".repeat(43))
            .and_then(|pairing| {
                pairing
                    .pending_binding
                    .as_ref()
                    .map(|pending| pending.id.clone())
            })
            .unwrap();
        json!({
            "version": 2,
            "type": match result {
                BindingResult::Ready => "office.binding_ready",
                BindingResult::Abort => "office.binding_abort",
                BindingResult::Committed => "office.binding_committed",
            },
            "pairing_id": "A".repeat(43),
            "binding_id": binding_id,
        })
        .as_object()
        .unwrap()
        .clone()
    }

    #[tokio::test]
    async fn saturated_diagnostic_response_never_marks_connection_failed() {
        let (sender, _receiver) = mpsc::channel(1);
        let failed = Arc::new(Notify::new());
        let tx = Tx {
            sender,
            failed: failed.clone(),
        };
        tx.sender
            .try_send(Message::Ping(Vec::new().into()))
            .unwrap();

        diagnostic_response(
            &tx,
            json!({"version":2,"type":"office.diagnostic.accepted","event_id":"event"}),
        );

        assert!(
            tokio::time::timeout(Duration::from_millis(10), failed.notified())
                .await
                .is_err()
        );
    }

    #[test]
    fn protocol_error_audit_excludes_payloads_and_identifiers() {
        let audit = protocol_error_log(
            42,
            false,
            r#"{"version":2,"type":"pc.chunk","session_id":"secret-session","request_id":"secret-request","data":"secret-document-data"}"#,
            "invalid_sequence",
        );
        assert_eq!(
            audit,
            json!({
                "event": "relay_protocol_error",
                "connection_id": 42,
                "role": "pc",
                "frame_type": "pc.chunk",
                "error_code": "invalid_sequence",
            })
        );
        let serialized = audit.to_string();
        assert!(!serialized.contains("secret-session"));
        assert!(!serialized.contains("secret-request"));
        assert!(!serialized.contains("secret-document-data"));

        let resume_audit = protocol_error_log(
            43,
            true,
            r#"{"version":2,"type":"office.prove","binding_id":"secret-binding","challenge":"secret-challenge","signature":"secret-signature","subject_hash":"secret-subject"}"#,
            "invalid_proof",
        )
        .to_string();
        for secret in [
            "secret-binding",
            "secret-challenge",
            "secret-signature",
            "secret-subject",
        ] {
            assert!(!resume_audit.contains(secret));
        }
    }

    #[tokio::test]
    async fn saturated_office_approval_queue_does_not_leave_a_live_binding() {
        let app = test_app();
        let (office_sender, _office_receiver) = mpsc::channel(1);
        office_sender
            .try_send(Message::Ping(Vec::new().into()))
            .unwrap();
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(2);
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let subject = [4_u8; 32];
        insert_test_pairing(&app, office_tx, pc_tx.clone(), subject).await;

        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Err("peer_unavailable")
        );
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
    }

    #[tokio::test]
    async fn saturated_pc_approval_queue_does_not_leave_a_live_binding() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(2);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(1);
        pc_sender
            .try_send(Message::Ping(Vec::new().into()))
            .unwrap();
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let subject = [6_u8; 32];
        insert_test_pairing(&app, office_tx, pc_tx.clone(), subject).await;

        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Ok(())
        );
        let _offer = office_receiver.recv().await.unwrap();
        let result = test_binding_result_frame(&app, BindingResult::Ready).await;
        assert_eq!(
            office_binding_result(&app, 1, result, BindingResult::Ready).await,
            Ok(())
        );
        let _commit = office_receiver.recv().await.unwrap();
        let result = test_binding_result_frame(&app, BindingResult::Committed).await;
        assert_eq!(
            office_binding_result(&app, 1, result, BindingResult::Committed).await,
            Err("peer_unavailable")
        );
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
    }

    #[tokio::test]
    async fn post_commit_delivery_failure_compensates_the_new_binding() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(2);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(2);
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let subject = [5_u8; 32];
        insert_test_pairing(&app, office_tx, pc_tx.clone(), subject).await;
        app.inner
            .fail_approval_delivery
            .store(true, Ordering::SeqCst);

        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Ok(())
        );
        let _offer = office_receiver.recv().await.unwrap();
        let result = test_binding_result_frame(&app, BindingResult::Ready).await;
        assert_eq!(
            office_binding_result(&app, 1, result, BindingResult::Ready).await,
            Ok(())
        );
        let _commit = office_receiver.recv().await.unwrap();
        let result = test_binding_result_frame(&app, BindingResult::Committed).await;
        assert_eq!(
            office_binding_result(&app, 1, result, BindingResult::Committed).await,
            Err("peer_unavailable")
        );
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
    }

    #[tokio::test]
    async fn failed_compensation_denies_resume_until_the_sweeper_converges_revocation() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(4);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(4);
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let subject = [9_u8; 32];
        insert_test_pairing(&app, office_tx.clone(), pc_tx.clone(), subject).await;
        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Ok(())
        );
        let _offer = office_receiver.recv().await.unwrap();
        let ready = test_binding_result_frame(&app, BindingResult::Ready).await;
        assert_eq!(
            office_binding_result(&app, 1, ready, BindingResult::Ready).await,
            Ok(())
        );
        let _commit = office_receiver.recv().await.unwrap();
        let committed = test_binding_result_frame(&app, BindingResult::Committed).await;
        let binding_id = string(&committed, "binding_id").unwrap().to_owned();
        app.inner
            .fail_approval_delivery
            .store(true, Ordering::SeqCst);
        app.inner.bindings.fail_next_revoke();
        assert_eq!(
            office_binding_result(&app, 1, committed, BindingResult::Committed).await,
            Err("peer_unavailable")
        );
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
        {
            let store = app.inner.state.lock().await;
            assert!(store.denied_bindings.contains(&binding_id));
            assert!(store.pending_revocations.contains_key(&binding_id));
            assert!(!store.codes.contains_key("123456"));
        }
        let resume = json!({
            "version": 2,
            "type": "office.resume",
            "binding_id": binding_id,
            "host": "Word",
            "capabilities": ["agent.v1"],
        })
        .as_object()
        .unwrap()
        .clone();
        assert_eq!(
            office_resume(
                &app,
                3,
                &office_tx,
                Some(OFFICE_ORIGIN),
                "2001:db8::9".parse().unwrap(),
                resume,
            )
            .await,
            Err("binding_unavailable")
        );
        {
            let mut store = app.inner.state.lock().await;
            expire(&app.inner, &mut store);
            assert!(!store.denied_bindings.contains(&binding_id));
            assert!(!store.pending_revocations.contains_key(&binding_id));
        }
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
    }

    #[tokio::test]
    async fn approval_delivery_failure_cannot_resume_after_relay_restart() {
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "wiswork-relay-failed-activation-{}-{unique}.sqlite",
            std::process::id()
        ));
        let app =
            test_app_with_bindings(Config::default(), BindingStore::open(Some(&path)).unwrap());
        let (office_sender, mut office_receiver) = mpsc::channel(4);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(4);
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let subject = [11_u8; 32];
        insert_test_pairing(&app, office_tx, pc_tx.clone(), subject).await;
        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Ok(())
        );
        let _offer = office_receiver.recv().await.unwrap();
        let ready = test_binding_result_frame(&app, BindingResult::Ready).await;
        assert_eq!(
            office_binding_result(&app, 1, ready, BindingResult::Ready).await,
            Ok(())
        );
        let _commit = office_receiver.recv().await.unwrap();
        let committed = test_binding_result_frame(&app, BindingResult::Committed).await;
        let binding_id = string(&committed, "binding_id").unwrap().to_owned();
        app.inner
            .fail_approval_delivery
            .store(true, Ordering::SeqCst);
        app.inner.bindings.fail_next_revoke();
        assert_eq!(
            office_binding_result(&app, 1, committed, BindingResult::Committed).await,
            Err("peer_unavailable")
        );
        drop(app);

        let restarted =
            test_app_with_bindings(Config::default(), BindingStore::open(Some(&path)).unwrap());
        let (retry_sender, _retry_receiver) = mpsc::channel(2);
        let retry_tx = Tx {
            sender: retry_sender,
            failed: Arc::new(Notify::new()),
        };
        let resume = json!({
            "version": 2,
            "type": "office.resume",
            "binding_id": binding_id,
            "host": "Word",
            "capabilities": ["agent.v1"],
        })
        .as_object()
        .unwrap()
        .clone();
        assert_eq!(
            office_resume(
                &restarted,
                3,
                &retry_tx,
                Some(OFFICE_ORIGIN),
                "2001:db8::11".parse().unwrap(),
                resume,
            )
            .await,
            Err("binding_unavailable")
        );
        drop(restarted);
        std::fs::remove_file(path).unwrap();
    }

    #[tokio::test]
    async fn session_capacity_failure_compensates_and_terminates_both_approved_peers() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(4);
        let office_failed = Arc::new(Notify::new());
        let office_tx = Tx {
            sender: office_sender,
            failed: office_failed.clone(),
        };
        let (pc_sender, _pc_receiver) = mpsc::channel(4);
        let pc_failed = Arc::new(Notify::new());
        let pc_tx = Tx {
            sender: pc_sender,
            failed: pc_failed.clone(),
        };
        let subject = [10_u8; 32];
        insert_test_pairing(&app, office_tx.clone(), pc_tx.clone(), subject).await;
        assert_eq!(
            approve(&app, 2, &pc_tx, test_approval_frame()).await,
            Ok(())
        );
        let _offer = office_receiver.recv().await.unwrap();
        let ready = test_binding_result_frame(&app, BindingResult::Ready).await;
        assert_eq!(
            office_binding_result(&app, 1, ready, BindingResult::Ready).await,
            Ok(())
        );
        let _commit = office_receiver.recv().await.unwrap();
        {
            let mut store = app.inner.state.lock().await;
            let now = Instant::now();
            for index in 0..10_000 {
                store.sessions.insert(
                    format!("dummy-{index}"),
                    Session {
                        version: PROTOCOL_V2,
                        host: "Word".to_owned(),
                        office: 100,
                        office_tx: office_tx.clone(),
                        pc: 101,
                        pc_tx: pc_tx.clone(),
                        office_cap: "office".to_owned(),
                        pc_cap: "pc".to_owned(),
                        expires: now + Duration::from_secs(60),
                        absolute_expires: now + Duration::from_secs(60),
                        active: None,
                        used_requests: VecDeque::new(),
                        capabilities: vec!["agent.v1".to_owned()],
                        diagnostics: 0,
                        diagnostic_window_started: now,
                        diagnostic_window_count: 0,
                        binding_id: None,
                    },
                );
            }
        }
        let committed = test_binding_result_frame(&app, BindingResult::Committed).await;
        assert_eq!(
            office_binding_result(&app, 1, committed, BindingResult::Committed).await,
            Err("relay_busy")
        );
        assert_eq!(app.inner.bindings.live_count(&subject), 0);
        {
            let store = app.inner.state.lock().await;
            assert!(!store.codes.contains_key("123456"));
        }
        tokio::time::timeout(Duration::from_millis(10), office_failed.notified())
            .await
            .expect("Office writer must terminate");
        tokio::time::timeout(Duration::from_millis(10), pc_failed.notified())
            .await
            .expect("PC writer must terminate");
    }

    #[test]
    fn challenge_sweeper_removes_entries_at_the_proof_deadline() {
        let mut store = Store::default();
        store.resume_challenges.insert(
            1,
            ResumeChallenge {
                binding: Binding {
                    id: "A".repeat(43),
                    subject: [1; 32],
                    public_key: [0; 65],
                    host: "Word".to_owned(),
                    origin: OFFICE_ORIGIN.to_owned(),
                    capabilities: vec!["agent.v1".to_owned()],
                },
                challenge: "challenge".to_owned(),
                expires: Instant::now(),
            },
        );

        let app = test_app();
        expire(&app.inner, &mut store);

        assert!(store.resume_challenges.is_empty());
        assert!(store.expired_resume_challenges.contains_key(&1));
    }

    #[tokio::test]
    async fn swept_challenge_still_reports_challenge_expired_to_its_connection() {
        let app = test_app();
        app.inner
            .state
            .lock()
            .await
            .expired_resume_challenges
            .insert(
                1,
                ExpiredResumeChallenge {
                    binding_hash: Sha256::digest("A".repeat(43).as_bytes()).into(),
                    challenge_hash: Sha256::digest(b"challenge").into(),
                    expires: Instant::now() + Duration::from_secs(60),
                },
            );
        let (sender, _receiver) = mpsc::channel(1);
        let tx = Tx {
            sender,
            failed: Arc::new(Notify::new()),
        };
        let frame = json!({
            "version": 2,
            "type": "office.prove",
            "binding_id": "A".repeat(43),
            "challenge": "challenge",
            "signature": "signature",
        })
        .as_object()
        .unwrap()
        .clone();
        let mut unrelated = frame.clone();
        unrelated.insert("challenge".to_owned(), Value::String("other".to_owned()));
        assert_eq!(
            office_prove(&app, 1, &tx, unrelated).await,
            Err("invalid_proof")
        );
        assert_eq!(
            office_prove(&app, 1, &tx, frame).await,
            Err("challenge_expired")
        );
    }

    #[tokio::test]
    async fn sweeper_notifies_both_kinds_of_expired_resume_waiter() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(1);
        let (pc_sender, mut pc_receiver) = mpsc::channel(1);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        let binding = Binding {
            id: "A".repeat(43),
            subject: [1; 32],
            public_key: [0; 65],
            host: "Word".to_owned(),
            origin: OFFICE_ORIGIN.to_owned(),
            capabilities: vec!["agent.v1".to_owned()],
        };
        let mut store = Store::default();
        store.office_resumes.insert(
            1,
            OfficeResume {
                binding,
                office: 1,
                office_tx,
                expires: Instant::now(),
            },
        );
        store.pc_resumes.insert(
            2,
            PcResume {
                binding_id: "A".repeat(43),
                pc: 2,
                pc_tx,
                subject: [1; 32],
                capabilities: vec!["agent.v1".to_owned()],
                expires: Instant::now(),
            },
        );
        expire(&app.inner, &mut store);
        for receiver in [&mut office_receiver, &mut pc_receiver] {
            let Message::Text(text) = receiver.try_recv().unwrap() else {
                panic!("expected text frame")
            };
            let frame: Value = serde_json::from_str(&text).unwrap();
            assert_eq!(frame["code"], "peer_unavailable");
        }
    }

    #[tokio::test]
    async fn expired_binding_offer_aborts_and_degrades_to_a_short_session() {
        let app = test_app();
        let (office_sender, mut office_receiver) = mpsc::channel(4);
        let (pc_sender, mut pc_receiver) = mpsc::channel(4);
        let office_tx = Tx {
            sender: office_sender,
            failed: Arc::new(Notify::new()),
        };
        let pc_tx = Tx {
            sender: pc_sender,
            failed: Arc::new(Notify::new()),
        };
        insert_test_pairing(&app, office_tx, pc_tx, [8; 32]).await;
        {
            let mut store = app.inner.state.lock().await;
            let pairing = store.pairings.get_mut(&"A".repeat(43)).unwrap();
            let (pc, pc_tx) = pairing.pc.clone().unwrap();
            pairing.pending_binding = Some(PendingBinding {
                id: "A".repeat(43),
                pc,
                pc_tx,
                subject: pairing.pc_subject.unwrap(),
                capabilities: pairing.negotiated_capabilities.clone(),
                phase: PendingBindingPhase::Offered,
            });
            pairing.expires = Instant::now();
        }
        let late_ready = test_binding_result_frame(&app, BindingResult::Ready).await;
        let duplicate_late_ready = late_ready.clone();
        assert_eq!(
            office_binding_result(&app, 1, late_ready, BindingResult::Ready).await,
            Ok(())
        );
        let Message::Text(aborted) = office_receiver.recv().await.unwrap() else {
            panic!("expected binding abort")
        };
        let aborted: Value = serde_json::from_str(&aborted).unwrap();
        assert_eq!(aborted["type"], "office.binding_aborted");
        let Message::Text(office_approved) = office_receiver.recv().await.unwrap() else {
            panic!("expected office approval")
        };
        let Message::Text(pc_approved) = pc_receiver.recv().await.unwrap() else {
            panic!("expected pc approval")
        };
        for text in [office_approved, pc_approved] {
            let frame: Value = serde_json::from_str(&text).unwrap();
            assert_eq!(frame["features"], json!([]));
            assert!(frame.get("binding_id").is_none());
        }
        assert_eq!(
            office_binding_result(&app, 1, duplicate_late_ready, BindingResult::Ready,).await,
            Ok(())
        );
        assert_eq!(app.inner.state.lock().await.sessions.len(), 1);
    }

    #[test]
    fn resume_attempt_ip_guard_does_not_grow_past_ten_thousand() {
        let mut store = Store::default();
        for index in 0..10_000_u128 {
            store.resume_attempts.insert(
                IpAddr::V6(std::net::Ipv6Addr::from(index)),
                (Instant::now(), 1),
            );
        }

        assert_eq!(
            consume_office_resume_attempt(
                &mut store,
                20_001,
                IpAddr::V6(std::net::Ipv6Addr::from(10_001_u128)),
                Duration::from_secs(120),
                20_000,
            ),
            Err("relay_busy")
        );
        assert_eq!(store.resume_attempts.len(), 10_000);
        assert!(!store.connection_resume_attempts.contains_key(&20_001));
    }

    #[tokio::test]
    async fn global_resume_budget_precedes_binding_lookup_and_ip_tracking() {
        let config = Config {
            max_global_resume_attempts: 1,
            ..Config::default()
        };
        let app = test_app_with_config(config);
        let (sender, _receiver) = mpsc::channel(4);
        let tx = Tx {
            sender,
            failed: Arc::new(Notify::new()),
        };
        let frame = || {
            json!({
                "version": 2,
                "type": "office.resume",
                "binding_id": "A".repeat(43),
                "host": "Word",
                "capabilities": ["agent.v1"]
            })
            .as_object()
            .unwrap()
            .clone()
        };

        assert_eq!(
            office_resume(
                &app,
                1,
                &tx,
                Some(OFFICE_ORIGIN),
                "2001:db8::1".parse().unwrap(),
                frame(),
            )
            .await,
            Err("binding_unavailable")
        );
        assert_eq!(app.inner.bindings.get_live_call_count(), 1);
        for (conn, ip) in [(2, "2001:db8::2"), (3, "2001:db8::3")] {
            assert_eq!(
                office_resume(
                    &app,
                    conn,
                    &tx,
                    Some(OFFICE_ORIGIN),
                    ip.parse().unwrap(),
                    frame(),
                )
                .await,
                Err("resume_rate_limited")
            );
        }
        assert_eq!(app.inner.bindings.get_live_call_count(), 1);
        let store = app.inner.state.lock().await;
        assert_eq!(store.resume_attempts.len(), 1);
        assert!(!store.connection_resume_attempts.contains_key(&2));
        assert!(!store.connection_resume_attempts.contains_key(&3));
    }

    #[test]
    fn expired_resume_attempt_window_clears_and_accepts_a_fresh_attempt() {
        let mut store = Store::default();
        let window = Duration::from_secs(1);
        store.global_resume_attempts = (Some(Instant::now() - window - window), 99);
        store.resume_attempts.insert(
            "2001:db8::1".parse().unwrap(),
            (Instant::now() - window - window, 20),
        );

        assert_eq!(store.resume_attempts.len(), 1);
        assert_eq!(
            consume_office_resume_attempt(&mut store, 1, "2001:db8::2".parse().unwrap(), window, 1,),
            Ok(())
        );
        assert_eq!(store.global_resume_attempts.1, 1);
        assert_eq!(store.resume_attempts.len(), 1);
        assert!(
            store
                .resume_attempts
                .contains_key(&"2001:db8::2".parse().unwrap())
        );
    }

    #[tokio::test]
    async fn sweeper_releases_the_app_and_binding_database() {
        let path = std::env::temp_dir().join(format!("wiswork-relay-sweeper-{}.sqlite3", token()));
        let config = Config {
            binding_database: Some(path.clone()),
            ..Config::default()
        };
        let app = test_app_with_config(config);
        let weak = Arc::downgrade(&app.inner);
        let sweeper = spawn_sweeper(&app);

        drop(app);
        tokio::time::timeout(Duration::from_secs(1), sweeper)
            .await
            .expect("sweeper should stop when the app is dropped")
            .unwrap();
        assert!(weak.upgrade().is_none());
        drop(BindingStore::open(Some(&path)).unwrap());
        std::fs::remove_file(path).unwrap();
    }
}
