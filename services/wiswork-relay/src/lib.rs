use axum::{
    Router,
    extract::{
        State,
        ws::{Message, WebSocket, WebSocketUpgrade},
    },
    http::{HeaderMap, StatusCode},
    response::IntoResponse,
    routing::get,
};
use base64::{Engine as _, engine::general_purpose::STANDARD};
use futures_util::{SinkExt, StreamExt};
use rand::{Rng, distr::Alphanumeric};
use serde_json::{Map, Value, json};
use std::{
    collections::{HashMap, HashSet},
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};
use tokio::sync::{Mutex, mpsc};

pub const OFFICE_ORIGIN: &str = "https://office.8-216-134-194.sslip.io";
const CONTROL_MAX: usize = 16 * 1024;
const REQUEST_MAX: usize = 256 * 1024;
const CHUNK_MAX: usize = 64 * 1024;
const RESPONSE_MAX: usize = 16 * 1024 * 1024;

#[derive(Clone)]
pub struct Config {
    pub pairing_ttl: Duration,
    pub session_ttl: Duration,
    pub request_ttl: Duration,
    pub max_claim_attempts: u8,
}
impl Default for Config {
    fn default() -> Self {
        Self {
            pairing_ttl: Duration::from_secs(120),
            session_ttl: Duration::from_secs(1800),
            request_ttl: Duration::from_secs(120),
            max_claim_attempts: 5,
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
}
type Tx = mpsc::Sender<Message>;
struct Pairing {
    id: String,
    code: String,
    host: String,
    office: u64,
    office_tx: Tx,
    pc: Option<(u64, Tx)>,
    expires: Instant,
    attempts: u8,
}
struct Active {
    id: String,
    sequence: u64,
    bytes: usize,
    deadline: Instant,
    started: bool,
}
struct Session {
    office: u64,
    office_tx: Tx,
    pc: u64,
    pc_tx: Tx,
    office_cap: String,
    pc_cap: String,
    expires: Instant,
    active: Option<Active>,
    used_requests: HashSet<String>,
}
#[derive(Default)]
struct Store {
    pairings: HashMap<String, Pairing>,
    codes: HashMap<String, String>,
    sessions: HashMap<String, Session>,
    claim_attempts: HashMap<u64, u8>,
}

pub fn app(config: Config) -> Router {
    let state = App {
        inner: Arc::new(Inner {
            state: Mutex::new(Store::default()),
            next: AtomicU64::new(1),
            config,
        }),
    };
    let sweeper = state.clone();
    tokio::spawn(async move {
        let mut interval = tokio::time::interval(Duration::from_secs(1));
        loop {
            interval.tick().await;
            let mut store = sweeper.inner.state.lock().await;
            expire(&mut store);
        }
    });
    Router::new()
        .route("/office-relay", get(upgrade))
        .route("/office-relay/health", get(|| async { "ok" }))
        .with_state(state)
}

async fn upgrade(
    State(app): State<App>,
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
    ws.max_message_size(CONTROL_MAX.max(REQUEST_MAX))
        .on_upgrade(move |socket| connection(app, socket, origin))
        .into_response()
}

async fn connection(app: App, socket: WebSocket, origin: Option<String>) {
    let id = app.inner.next.fetch_add(1, Ordering::Relaxed);
    let (mut sink, mut stream) = socket.split();
    let (tx, mut rx) = mpsc::channel(64);
    let writer = tokio::spawn(async move {
        while let Some(message) = rx.recv().await {
            if sink.send(message).await.is_err() {
                break;
            }
        }
    });
    while let Some(Ok(message)) = stream.next().await {
        match message {
            Message::Text(text) if text.len() <= REQUEST_MAX => {
                if let Err(code) = process(&app, id, &tx, origin.as_deref(), text.as_str()).await {
                    error(&tx, code);
                    break;
                }
            }
            Message::Ping(data) => {
                let _ = tx.try_send(Message::Pong(data));
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
    let _ = tx.try_send(Message::Text(value.to_string().into()));
}
fn error(tx: &Tx, code: &str) {
    send(tx, json!({"version":1,"type":"relay.error","code":code}));
}
fn valid_base(m: &Map<String, Value>) -> bool {
    m.get("version").and_then(Value::as_u64) == Some(1)
        && m.get("type").and_then(Value::as_str).is_some()
}

async fn process(
    app: &App,
    conn: u64,
    tx: &Tx,
    origin: Option<&str>,
    text: &str,
) -> Result<(), &'static str> {
    let map = object(text, REQUEST_MAX)?;
    if !valid_base(&map) {
        return Err("invalid_frame");
    }
    let kind = string(&map, "type")?;
    if text.len() > CONTROL_MAX && !matches!(kind, "office.request" | "pc.chunk") {
        return Err("frame_too_large");
    }
    if origin.is_some() && !kind.starts_with("office.") {
        return Err("role_not_allowed");
    }
    if origin.is_none() && !kind.starts_with("pc.") {
        return Err("role_not_allowed");
    }
    match kind {
        "office.create" => create(app, conn, tx, origin, map).await,
        "pc.claim" => claim(app, conn, tx, map).await,
        "pc.approve" => approve(app, conn, tx, map).await,
        "pc.reject" => reject(app, conn, map).await,
        "office.request" => request(app, conn, map, text.len()).await,
        "office.cancel" => cancel(app, conn, map).await,
        "pc.chunk" => chunk(app, conn, map).await,
        "pc.start" => start(app, conn, map).await,
        "pc.done" => done(app, conn, map).await,
        "pc.error" => pc_error(app, conn, map).await,
        _ => Err("unknown_type"),
    }
}

async fn create(
    app: &App,
    conn: u64,
    tx: &Tx,
    origin: Option<&str>,
    m: Map<String, Value>,
) -> Result<(), &'static str> {
    if !exact(&m, &["version", "type", "host"]) || origin != Some(OFFICE_ORIGIN) {
        return Err("invalid_frame");
    }
    let host = string(&m, "host")?;
    if !matches!(host, "Word" | "Excel" | "PowerPoint") {
        return Err("unsupported_host");
    }
    let mut store = app.inner.state.lock().await;
    expire(&mut store);
    if store.pairings.len() >= 10_000 {
        return Err("relay_busy");
    }
    let id = token();
    let secret = token();
    let mut rng = rand::rng();
    let mut code = format!("{:06}", rng.random_range(0..1_000_000));
    while store.codes.contains_key(&code) {
        code = format!("{:06}", rng.random_range(0..1_000_000));
    }
    let pairing = Pairing {
        id: id.clone(),
        code: code.clone(),
        host: host.into(),
        office: conn,
        office_tx: tx.clone(),
        pc: None,
        expires: Instant::now() + app.inner.config.pairing_ttl,
        attempts: 0,
    };
    store.codes.insert(code.clone(), id.clone());
    store.pairings.insert(id.clone(), pairing);
    send(
        tx,
        json!({"version":1,"type":"office.created","pairing_id":id,"polling_secret":secret,"verification_code":code,"expires_in":app.inner.config.pairing_ttl.as_secs()}),
    );
    Ok(())
}

async fn claim(app: &App, conn: u64, tx: &Tx, m: Map<String, Value>) -> Result<(), &'static str> {
    if !exact(&m, &["version", "type", "verification_code"]) {
        return Err("invalid_frame");
    }
    let code = string(&m, "verification_code")?;
    if code.len() != 6 || !code.bytes().all(|b| b.is_ascii_digit()) {
        return Err("invalid_code");
    }
    let mut s = app.inner.state.lock().await;
    let attempts = s.claim_attempts.entry(conn).or_default();
    *attempts = attempts.saturating_add(1);
    if *attempts > app.inner.config.max_claim_attempts {
        return Err("claim_limit");
    }
    expire(&mut s);
    let Some(id) = s.codes.get(code).cloned() else {
        return Err("invalid_code");
    };
    let max = app.inner.config.max_claim_attempts;
    let p = s.pairings.get_mut(&id).ok_or("invalid_code")?;
    p.attempts += 1;
    if p.attempts > max {
        return Err("claim_limit");
    };
    if p.pc.is_some() {
        return Err("already_claimed");
    }
    p.pc = Some((conn, tx.clone()));
    send(
        tx,
        json!({"version":1,"type":"pc.claimed","pairing_id":p.id,"host":p.host,"origin":OFFICE_ORIGIN,"verification_code":p.code,"expires_in":p.expires.saturating_duration_since(Instant::now()).as_secs()}),
    );
    Ok(())
}

async fn approve(app: &App, conn: u64, tx: &Tx, m: Map<String, Value>) -> Result<(), &'static str> {
    if !exact(&m, &["version", "type", "pairing_id"]) {
        return Err("invalid_frame");
    }
    let id = string(&m, "pairing_id")?.to_owned();
    let mut s = app.inner.state.lock().await;
    expire(&mut s);
    if s.sessions.len() >= 10_000 {
        return Err("relay_busy");
    }
    let p = s.pairings.remove(&id).ok_or("invalid_pairing")?;
    if p.pc.as_ref().map(|x| x.0) != Some(conn) {
        s.pairings.insert(id, p);
        return Err("invalid_pairing");
    }
    s.codes.remove(&p.code);
    let sid = token();
    let oc = token();
    let pc = token();
    let expires = app.inner.config.session_ttl;
    send(
        &p.office_tx,
        json!({"version":1,"type":"office.approved","session_id":sid,"capability":oc,"expires_in":expires.as_secs()}),
    );
    send(
        tx,
        json!({"version":1,"type":"pc.approved","session_id":sid,"capability":pc,"expires_in":expires.as_secs()}),
    );
    s.sessions.insert(
        sid,
        Session {
            office: p.office,
            office_tx: p.office_tx,
            pc: conn,
            pc_tx: tx.clone(),
            office_cap: oc,
            pc_cap: pc,
            expires: Instant::now() + expires,
            active: None,
            used_requests: HashSet::new(),
        },
    );
    Ok(())
}

async fn reject(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    if !exact(&m, &["version", "type", "pairing_id"]) {
        return Err("invalid_frame");
    }
    let id = string(&m, "pairing_id")?;
    let mut s = app.inner.state.lock().await;
    let p = s.pairings.get(id).ok_or("invalid_pairing")?;
    if p.pc.as_ref().map(|v| v.0) != Some(conn) {
        return Err("invalid_pairing");
    }
    let p = s.pairings.remove(id).unwrap();
    s.codes.remove(&p.code);
    send(&p.office_tx, json!({"version":1,"type":"office.rejected"}));
    Ok(())
}

fn session_fields(m: &Map<String, Value>) -> Result<(&str, &str, &str), &'static str> {
    Ok((
        string(m, "session_id")?,
        string(m, "capability")?,
        string(m, "request_id")?,
    ))
}
async fn request(
    app: &App,
    conn: u64,
    m: Map<String, Value>,
    size: usize,
) -> Result<(), &'static str> {
    if !exact(
        &m,
        &[
            "version",
            "type",
            "session_id",
            "capability",
            "request_id",
            "body",
        ],
    ) {
        return Err("invalid_frame");
    }
    if size > REQUEST_MAX {
        return Err("request_too_large");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    if rid.is_empty() || rid.len() > 128 {
        return Err("invalid_frame");
    }
    let mut st = app.inner.state.lock().await;
    expire(&mut st);
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap {
        return Err("invalid_capability");
    }
    if session.active.is_some() {
        return Err("request_active");
    }
    if !session.used_requests.insert(rid.to_owned()) {
        return Err("duplicate_request");
    }
    session.active = Some(Active {
        id: rid.into(),
        sequence: 0,
        bytes: 0,
        deadline: Instant::now() + app.inner.config.request_ttl,
        started: false,
    });
    send(
        &session.pc_tx,
        json!({"version":1,"type":"relay.request","session_id":sid,"request_id":rid,"body":m["body"]}),
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
            send(
                &session.pc_tx,
                json!({"version":1,"type":"relay.cancel","session_id":deadline_sid,"request_id":deadline_rid}),
            );
            send(
                &session.office_tx,
                json!({"version":1,"type":"relay.error","session_id":deadline_sid,"request_id":deadline_rid,"code":"request_timeout"}),
            );
        }
    });
    Ok(())
}
async fn cancel(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    if !exact(
        &m,
        &["version", "type", "session_id", "capability", "request_id"],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let mut st = app.inner.state.lock().await;
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.office != conn || session.office_cap != cap {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|a| a.id.as_str()) != Some(rid) {
        return Err("invalid_request");
    }
    session.active = None;
    send(
        &session.pc_tx,
        json!({"version":1,"type":"relay.cancel","session_id":sid,"request_id":rid}),
    );
    Ok(())
}

async fn chunk(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
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
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap {
        return Err("invalid_capability");
    }
    let active = session.active.as_mut().ok_or("invalid_request")?;
    if active.id != rid || !active.started || active.sequence != seq {
        return Err("invalid_sequence");
    }
    if active.deadline <= Instant::now() {
        return Err("request_timeout");
    }
    active.sequence += 1;
    active.bytes = active
        .bytes
        .checked_add(decoded.len())
        .ok_or("response_too_large")?;
    if active.bytes > RESPONSE_MAX {
        return Err("response_too_large");
    }
    send(
        &session.office_tx,
        json!({"version":1,"type":"relay.chunk","session_id":sid,"request_id":rid,"sequence":seq,"data":data}),
    );
    Ok(())
}
async fn start(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
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
        .filter(|v| (100..=599).contains(v))
        .ok_or("invalid_status")?;
    let content_type = string(&m, "content_type")?;
    if content_type.is_empty()
        || content_type.len() > 128
        || content_type.bytes().any(|b| !(0x20..=0x7e).contains(&b))
    {
        return Err("invalid_content_type");
    }
    let mut store = app.inner.state.lock().await;
    let session = store.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap {
        return Err("invalid_capability");
    }
    let active = session.active.as_mut().ok_or("invalid_request")?;
    if active.id != rid || active.started {
        return Err("invalid_request");
    }
    active.started = true;
    send(
        &session.office_tx,
        json!({"version":1,"type":"relay.start","session_id":sid,"request_id":rid,"status":status,"content_type":content_type}),
    );
    Ok(())
}
async fn done(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
    if !exact(
        &m,
        &["version", "type", "session_id", "capability", "request_id"],
    ) {
        return Err("invalid_frame");
    }
    let (sid, cap, rid) = session_fields(&m)?;
    let mut st = app.inner.state.lock().await;
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap {
        return Err("invalid_capability");
    }
    if !session
        .active
        .as_ref()
        .is_some_and(|a| a.id == rid && a.started)
    {
        return Err("invalid_request");
    }
    session.active = None;
    send(
        &session.office_tx,
        json!({"version":1,"type":"relay.done","session_id":sid,"request_id":rid}),
    );
    Ok(())
}
async fn pc_error(app: &App, conn: u64, m: Map<String, Value>) -> Result<(), &'static str> {
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
    let session = st.sessions.get_mut(sid).ok_or("invalid_session")?;
    if session.pc != conn || session.pc_cap != cap {
        return Err("invalid_capability");
    }
    if session.active.as_ref().map(|a| a.id.as_str()) != Some(rid) {
        return Err("invalid_request");
    }
    session.active = None;
    send(
        &session.office_tx,
        json!({"version":1,"type":"relay.error","session_id":sid,"request_id":rid,"code":code}),
    );
    Ok(())
}

fn expire(s: &mut Store) {
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
            send(&p.office_tx, json!({"version":1,"type":"office.expired"}));
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
            error(&x.office_tx, "session_expired");
            if let Some(active) = x.active {
                send(
                    &x.pc_tx,
                    json!({"version":1,"type":"relay.cancel","session_id":id,"request_id":active.id}),
                );
            }
        }
    }
}
async fn cleanup(app: &App, conn: u64) {
    let mut s = app.inner.state.lock().await;
    let pids: Vec<_> = s
        .pairings
        .iter()
        .filter(|(_, p)| p.office == conn || p.pc.as_ref().is_some_and(|v| v.0 == conn))
        .map(|(id, _)| id.clone())
        .collect();
    for id in pids {
        if let Some(p) = s.pairings.remove(&id) {
            s.codes.remove(&p.code);
            if p.office != conn {
                send(
                    &p.office_tx,
                    json!({"version":1,"type":"office.pc_offline"}),
                );
            }
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
                error(&x.office_tx, "session_revoked")
            }
            if x.pc != conn
                && let Some(a) = x.active
            {
                send(
                    &x.pc_tx,
                    json!({"version":1,"type":"relay.cancel","session_id":id,"request_id":a.id}),
                );
            }
        }
    }
}
