use std::{
    net::{IpAddr, Ipv4Addr, SocketAddr},
    path::PathBuf,
};
use tokio::net::TcpListener;
use wiswork_relay::{Config, try_app};

#[tokio::main]
async fn main() {
    let port = match std::env::var("WISWORK_RELAY_PORT") {
        Ok(value) => value.parse::<u16>().expect("invalid WISWORK_RELAY_PORT"),
        Err(std::env::VarError::NotPresent) => 43190,
        Err(_) => panic!("invalid WISWORK_RELAY_PORT"),
    };
    let binding_database = match std::env::var("WISWORK_RELAY_BINDING_DB") {
        Ok(value) if !value.is_empty() => {
            let path = PathBuf::from(value);
            assert!(
                path.is_absolute(),
                "WISWORK_RELAY_BINDING_DB must be absolute"
            );
            path
        }
        _ => panic!("WISWORK_RELAY_BINDING_DB is required"),
    };
    let pairing_resume_enabled = match std::env::var("WISWORK_RELAY_PAIRING_RESUME") {
        Ok(value) if value == "0" => false,
        Ok(value) if value == "1" => true,
        Ok(_) => panic!("WISWORK_RELAY_PAIRING_RESUME must be 0 or 1"),
        Err(std::env::VarError::NotPresent) => true,
        Err(_) => panic!("invalid WISWORK_RELAY_PAIRING_RESUME"),
    };
    let address = SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port);
    let listener = TcpListener::bind(address)
        .await
        .expect("relay loopback bind failed");
    axum::serve(
        listener,
        try_app(Config {
            binding_database: Some(binding_database),
            pairing_resume_enabled,
            ..Config::default()
        })
        .expect("binding database initialization failed")
        .into_make_service_with_connect_info::<SocketAddr>(),
    )
    .await
    .expect("relay server failed");
}
