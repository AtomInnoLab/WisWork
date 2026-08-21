use std::net::{IpAddr, Ipv4Addr, SocketAddr};
use tokio::net::TcpListener;
use wiswork_relay::{Config, app};

#[tokio::main]
async fn main() {
    let port = std::env::var("WISWORK_RELAY_PORT")
        .ok()
        .and_then(|v| v.parse::<u16>().ok())
        .unwrap_or(43190);
    let address = SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port);
    let listener = TcpListener::bind(address)
        .await
        .expect("relay loopback bind failed");
    axum::serve(listener, app(Config::default()))
        .await
        .expect("relay server failed");
}
