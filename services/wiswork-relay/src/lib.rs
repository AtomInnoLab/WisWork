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
use base64::{Engine as _, engine::general_purpose::STANDARD};
use futures_util::{SinkExt, StreamExt};
use jsonwebtoken::{Algorithm, DecodingKey, Validation, decode, decode_header, jwk::JwkSet};
use rand::{Rng, distr::Alphanumeric};
use serde::Deserialize;
use serde_json::{Map, Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::{HashMap, HashSet, VecDeque},
    net::{IpAddr, SocketAddr},
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};
use tokio::sync::{Mutex, Notify, Semaphore, mpsc};

pub const OFFICE_ORIGIN: &str = "https://office.8-216-134-194.sslip.io";
const CONTROL_MAX: usize = 16 * 1024;
const REQUEST_MAX: usize = 256 * 1024;
const FRAME_MAX: usize = REQUEST_MAX + CONTROL_MAX;
const CHUNK_MAX: usize = 64 * 1024;
const RESPONSE_MAX: usize = 16 * 1024 * 1024;
const DIAGNOSTIC_MAX: usize = 4 * 1024;
const DIAGNOSTIC_SESSION_MAX: u16 = 100;
const PROTOCOL_V2: u64 = 2;
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
    pub diagnostic_window: Duration,
    pub max_diagnostics_per_window: u8,
    pub auth_url: String,
    pub jwks_url: String,
    pub issuer: String,
    pub audience: String,
}
impl Default for Config {
    fn default() -> Self {
        Self {
            pairing_ttl: Duration::from_secs(120),
            session_ttl: Duration::from_secs(1800),
            session_max_ttl: Duration::from_secs(8 * 60 * 60),
            request_ttl: Duration::from_secs(120),
            max_claim_attempts: 5,
            max_global_claims: 1_000,
            diagnostic_window: Duration::from_secs(1),
            max_diagnostics_per_window: 10,
            auth_url: "https://auth.dev.wispaper.ai/oidc/me".into(),
            jwks_url: "https://auth.dev.wispaper.ai/oidc/jwks".into(),
            issuer: "https://auth.dev.wispaper.ai/oidc".into(),
            audience: "y3xpwx3ytskxf66p0wztm".into(),
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
}
struct Active {
    id: String,
    sequence: u64,
    bytes: usize,
    deadline: Instant,
    started: bool,
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
    preauth_attempts: HashMap<IpAddr, (Instant, u8)>,
    global_preauth: (Option<Instant>, u32),
}

pub fn app(config: Config) -> Router {
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
        }),
    };
    let sweeper = state.clone();
    tokio::spawn(async move {
        let mut interval = tokio::time::interval(Duration::from_secs(1));
        loop {
            interval.tick().await;
            let mut store = sweeper.inner.state.lock().await;
            expire(&mut store, sweeper.inner.config.pairing_ttl);
        }
    });
    Router::new()
        .route("/office-relay", get(upgrade))
        .route("/office-relay/health", get(|| async { "ok" }))
        .with_state(state)
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
    let allowed = url.as_str() == "https://auth.dev.wispaper.ai/oidc/me"
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
    let allowed = url.as_str() == "https://auth.dev.wispaper.ai/oidc/jwks"
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
                    error_for_text(&tx, text.as_str(), code);
                    break;
                }
            }
            Message::Ping(data) => {
                tx.sender
                    .try_send(Message::Pong(data))
                    .unwrap_or_else(|_| tx.failed.notify_one());
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
            _ => {}
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
fn versioned_error(tx: &Tx, version: u64, code: &str) {
    send(
        tx,
        json!({"version":version,"type":"relay.error","code":code}),
    );
}
fn renew_session(session: &mut Session, idle_ttl: Duration) {
    session.expires = (Instant::now() + idle_ttl).min(session.absolute_expires);
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
        && !matches!(kind, "office.request" | "office.diagnostic" | "pc.chunk")
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
        "pc.negotiate" => negotiate(app, tx, subject.ok_or("auth_required")?, map).await,
        "pc.claim" => claim(app, conn, tx, subject.ok_or("auth_required")?, map).await,
        "pc.approve" => approve(app, conn, tx, map).await,
        "pc.reject" => reject(app, conn, map).await,
        "office.request" => request(app, conn, map).await,
        "office.cancel" => cancel(app, conn, map).await,
        "office.diagnostic" => {
            if let Err(code) = diagnostic(app, conn, tx, map, text.len()).await {
                versioned_error(tx, PROTOCOL_V2, code);
            }
            Ok(())
        }
        "pc.chunk" => chunk(app, conn, map).await,
        "pc.start" => start(app, conn, map).await,
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
    if version(&m)? != PROTOCOL_V2
        || !exact(
            &m,
            &["version", "type", "verification_code", "capabilities"],
        )
    {
        return Err("invalid_frame");
    }
    let offered = capabilities(&m)?;
    let code = string(&m, "verification_code")?;
    if code.len() != 6 || !code.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err("invalid_code");
    }
    let mut store = app.inner.state.lock().await;
    expire(&mut store, app.inner.config.pairing_ttl);
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
    pairing.negotiated_subject = Some(subject);
    send(
        tx,
        json!({"version":PROTOCOL_V2,"type":"pc.negotiated","pairing_version":pairing.version,"capabilities":negotiated}),
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
    let expected: &[&str] = if protocol == PROTOCOL_V2 {
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
    let host = string(&m, "host")?;
    if !matches!(host, "Word" | "Excel" | "PowerPoint") {
        return Err("unsupported_host");
    }
    let mut store = app.inner.state.lock().await;
    expire(&mut store, app.inner.config.pairing_ttl);
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
    };
    store.codes.insert(code.clone(), id.clone());
    store.pairings.insert(id.clone(), pairing);
    send(
        tx,
        json!({"version":protocol,"type":"office.created","pairing_id":id,"verification_code":code,"expires_in":app.inner.config.pairing_ttl.as_secs()}),
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
    let expected: &[&str] = if protocol == PROTOCOL_V2 {
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
    let code = string(&m, "verification_code")?;
    if code.len() != 6 || !code.bytes().all(|b| b.is_ascii_digit()) {
        return Err("invalid_code");
    }
    let mut s = app.inner.state.lock().await;
    expire(&mut s, app.inner.config.pairing_ttl);
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
    p.negotiated_subject = None;
    send(
        tx,
        if protocol == PROTOCOL_V2 {
            json!({"version":protocol,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs(),"capabilities":p.negotiated_capabilities})
        } else {
            json!({"version":1,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs()})
        },
    );
    Ok(())
}

async fn approve(app: &App, conn: u64, tx: &Tx, m: Map<String, Value>) -> Result<(), &'static str> {
    let protocol = version(&m)?;
    let expected: &[&str] = if protocol == PROTOCOL_V2 {
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
    let mut s = app.inner.state.lock().await;
    expire(&mut s, app.inner.config.pairing_ttl);
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
    let p = s.pairings.remove(&id).ok_or("invalid_pairing")?;
    let sid = token();
    let oc = token();
    let pc = token();
    let expires = app
        .inner
        .config
        .session_ttl
        .min(app.inner.config.session_max_ttl);
    let now = Instant::now();
    let pc_approved = if protocol == PROTOCOL_V2 {
        json!({"version":protocol,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs(),"capabilities":approved_capabilities})
    } else {
        json!({"version":1,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs()})
    };
    let office_approved = if protocol == PROTOCOL_V2 {
        json!({"version":protocol,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs(),"capabilities":approved_capabilities})
    } else {
        json!({"version":1,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs()})
    };
    if tx.sender.capacity() == 0
        || p.office_tx.sender.capacity() == 0
        || try_send(tx, pc_approved).is_err()
        || try_send(&p.office_tx, office_approved).is_err()
    {
        tx.failed.notify_one();
        p.office_tx.failed.notify_one();
        return Err("peer_unavailable");
    }
    s.codes.remove(&p.code);
    s.sessions.insert(
        sid,
        Session {
            version: protocol,
            host: p.host,
            office: p.office,
            office_tx: p.office_tx,
            pc: conn,
            pc_tx: tx.clone(),
            office_cap: oc,
            pc_cap: pc,
            expires: now + expires,
            absolute_expires: now + app.inner.config.session_max_ttl,
            active: None,
            used_requests: VecDeque::new(),
            capabilities: approved_capabilities,
            diagnostics: 0,
            diagnostic_window_started: now,
            diagnostic_window_count: 0,
        },
    );
    Ok(())
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
    expire(&mut store, app.inner.config.pairing_ttl);
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
    send(
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
    expire(&mut st, app.inner.config.pairing_ttl);
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
    expire(&mut st, app.inner.config.pairing_ttl);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|a| a.id.as_str()) != Some(rid) {
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
    expire(&mut st, app.inner.config.pairing_ttl);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
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
    expire(&mut store, app.inner.config.pairing_ttl);
    let session = store.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
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
    expire(&mut st, app.inner.config.pairing_ttl);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if !session
        .active
        .as_ref()
        .is_some_and(|a| a.id == rid && a.started)
    {
        return Err("invalid_request");
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
    expire(&mut st, app.inner.config.pairing_ttl);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap || session.version != protocol {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|a| a.id.as_str()) != Some(rid) {
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

fn expire(s: &mut Store, ttl: Duration) {
    let now = Instant::now();
    let dead: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, p)| p.expires <= now)
        .map(|(id, _)| id.clone())
        .collect();
    for id in dead {
        if let Some(p) = s.pairings.remove(&id) {
            s.codes.remove(&p.code);
            send(
                &p.office_tx,
                json!({"version":p.version,"type":"office.expired"}),
            );
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
            versioned_error(&x.office_tx, x.version, "session_expired");
            if let Some(active) = x.active {
                send(
                    &x.pc_tx,
                    json!({"version":x.version,"type":"relay.cancel","session_id":id,"request_id":active.id}),
                );
            }
        }
    }
    s.claim_attempts
        .retain(|_, (started, _)| started.elapsed() <= ttl);
    s.create_attempts
        .retain(|_, (started, _)| started.elapsed() <= ttl);
    s.preauth_attempts
        .retain(|_, (started, _)| started.elapsed() <= ttl);
}
async fn cleanup(app: &App, conn: u64) {
    let mut s = app.inner.state.lock().await;
    s.connection_pairings.remove(&conn);
    let pids: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, p)| p.office == conn)
        .map(|(id, _)| id.clone())
        .collect();
    for id in pids {
        if let Some(p) = s.pairings.remove(&id) {
            s.codes.remove(&p.code);
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
