use std::net::{IpAddr, Ipv4Addr, SocketAddr};
use tokio::net::TcpListener;
use wiswork_relay::{Config, app};

#[tokio::main]
async fn main() {
    let port = match std::env::var("WISWORK_RELAY_PORT") {
        Ok(value) => value.parse::<u16>().expect("invalid WISWORK_RELAY_PORT"),
        Err(std::env::VarError::NotPresent) => 43190,
        Err(_) => panic!("invalid WISWORK_RELAY_PORT"),
    };
    let address = SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port);
    let listener = TcpListener::bind(address)
        .await
        .expect("relay loopback bind failed");
    axum::serve(
        listener,
        app(Config::default()).into_make_service_with_connect_info::<SocketAddr>(),
    )
    .await
    .expect("relay server failed");
}
