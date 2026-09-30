use std::io::{self, Write};
use std::net::SocketAddr;

use axum::{Json, Router, http::HeaderMap, routing::get};
use serde_json::json;
use tokio::net::TcpListener;
use wiswork_relay::{Config, app};

// Loopback-only fixture for exercising the real Relay with the Office and PC clients.
#[tokio::main]
async fn main() {
    let auth = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let auth_address = auth.local_addr().unwrap();
    let auth_app = Router::new().route(
        "/oidc/me",
        get(|headers: HeaderMap| async move {
            if headers
                .get("authorization")
                .and_then(|value| value.to_str().ok())
                == Some("Bearer local-business-smoke-token")
            {
                Ok(Json(json!({ "sub": "local-business-smoke-user" })))
            } else {
                Err(axum::http::StatusCode::UNAUTHORIZED)
            }
        }),
    );
    tokio::spawn(async move { axum::serve(auth, auth_app).await.unwrap() });

    let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
    let address: SocketAddr = listener.local_addr().unwrap();
    let config = Config {
        auth_url: format!("http://{auth_address}/oidc/me"),
        // This fixture exercises several valid sessions from one test account.
        max_claim_attempts: 12,
        ..Config::default()
    };
    println!("http://{address}");
    io::stdout().flush().unwrap();
    axum::serve(
        listener,
        app(config).into_make_service_with_connect_info::<SocketAddr>(),
    )
    .await
    .unwrap();
}
