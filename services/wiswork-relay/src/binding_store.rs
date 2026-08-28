use rusqlite::{Connection, OptionalExtension, TransactionBehavior, params};
use std::{
    fmt,
    path::Path,
    sync::Mutex,
    time::{SystemTime, UNIX_EPOCH},
};

const SCHEMA_VERSION: i64 = 1;
const MAX_BINDINGS_PER_SUBJECT: i64 = 12;
const MAX_REVOKED_BINDINGS_PER_SUBJECT: i64 = 24;
const PENDING_BINDING: i64 = -1;

#[derive(Debug)]
pub enum BindingStoreError {
    Database(rusqlite::Error),
    Io(std::io::Error),
    FutureSchema(i64),
    BindingLimit,
    CorruptRecord,
}

impl fmt::Display for BindingStoreError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Database(_) => formatter.write_str("binding database error"),
            Self::Io(_) => formatter.write_str("binding database file error"),
            Self::FutureSchema(version) => {
                write!(
                    formatter,
                    "unsupported binding database schema version {version}"
                )
            }
            Self::BindingLimit => formatter.write_str("binding limit reached"),
            Self::CorruptRecord => formatter.write_str("corrupt binding database record"),
        }
    }
}

impl std::error::Error for BindingStoreError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Database(error) => Some(error),
            Self::Io(error) => Some(error),
            _ => None,
        }
    }
}

impl From<rusqlite::Error> for BindingStoreError {
    fn from(value: rusqlite::Error) -> Self {
        Self::Database(value)
    }
}

impl From<std::io::Error> for BindingStoreError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

#[derive(Clone)]
pub(crate) struct Binding {
    pub id: String,
    pub subject: [u8; 32],
    pub public_key: [u8; 65],
    pub host: String,
    pub origin: String,
    pub capabilities: Vec<String>,
}

pub(crate) struct BindingStore {
    connection: Option<Mutex<Connection>>,
    #[cfg(test)]
    get_live_calls: std::sync::atomic::AtomicUsize,
    #[cfg(test)]
    fail_next_revoke: std::sync::atomic::AtomicBool,
}

impl BindingStore {
    pub fn disabled() -> Self {
        Self {
            connection: None,
            #[cfg(test)]
            get_live_calls: std::sync::atomic::AtomicUsize::new(0),
            #[cfg(test)]
            fail_next_revoke: std::sync::atomic::AtomicBool::new(false),
        }
    }

    pub fn open(path: Option<&Path>) -> Result<Self, BindingStoreError> {
        let connection = match path {
            Some(path) => {
                precreate_private_file(path)?;
                Connection::open(path)?
            }
            None => Connection::open_in_memory()?,
        };
        connection.busy_timeout(std::time::Duration::from_secs(2))?;
        connection.execute_batch("PRAGMA foreign_keys = ON; PRAGMA secure_delete = ON;")?;
        let store = Self {
            connection: Some(Mutex::new(connection)),
            #[cfg(test)]
            get_live_calls: std::sync::atomic::AtomicUsize::new(0),
            #[cfg(test)]
            fail_next_revoke: std::sync::atomic::AtomicBool::new(false),
        };
        store.initialize()?;
        Ok(store)
    }

    fn initialize(&self) -> Result<(), BindingStoreError> {
        let mut connection = self
            .connection
            .as_ref()
            .expect("enabled binding store")
            .lock()
            .expect("binding database mutex");
        let transaction = connection.transaction_with_behavior(TransactionBehavior::Immediate)?;
        let version: i64 = transaction.query_row("PRAGMA user_version", [], |row| row.get(0))?;
        match version {
            0 => transaction.execute_batch(
                "CREATE TABLE durable_bindings (
                    binding_id TEXT PRIMARY KEY NOT NULL,
                    subject_hash BLOB NOT NULL CHECK(length(subject_hash) = 32),
                    public_key BLOB NOT NULL CHECK(length(public_key) = 65),
                    host TEXT NOT NULL CHECK(host IN ('Word', 'Excel', 'PowerPoint')),
                    origin TEXT NOT NULL,
                    capabilities_json TEXT NOT NULL,
                    created_at INTEGER NOT NULL,
                    last_used_at INTEGER NOT NULL,
                    revoked_at INTEGER
                 ) STRICT;
                 CREATE INDEX durable_bindings_live_subject
                    ON durable_bindings(subject_hash) WHERE revoked_at IS NULL;
                 PRAGMA user_version = 1;",
            )?,
            SCHEMA_VERSION => {
                transaction.prepare(
                    "SELECT binding_id, subject_hash, public_key, host, origin,
                            capabilities_json, created_at, last_used_at, revoked_at
                     FROM durable_bindings LIMIT 0",
                )?;
                transaction.execute(
                    "DELETE FROM durable_bindings WHERE revoked_at = ?1",
                    [PENDING_BINDING],
                )?;
            }
            future => return Err(BindingStoreError::FutureSchema(future)),
        }
        transaction.commit()?;
        Ok(())
    }

    #[cfg(test)]
    pub fn enroll(&self, binding: &Binding) -> Result<(), BindingStoreError> {
        self.enroll_with_state(binding, None)
    }

    pub fn enroll_pending(&self, binding: &Binding) -> Result<(), BindingStoreError> {
        self.enroll_with_state(binding, Some(PENDING_BINDING))
    }

    fn enroll_with_state(
        &self,
        binding: &Binding,
        revoked_at: Option<i64>,
    ) -> Result<(), BindingStoreError> {
        let now = unix_time()?;
        let capabilities = serde_json::to_string(&binding.capabilities)
            .map_err(|_| BindingStoreError::CorruptRecord)?;
        let mut connection = self
            .connection
            .as_ref()
            .ok_or(BindingStoreError::CorruptRecord)?
            .lock()
            .expect("binding database mutex");
        let transaction = connection.transaction_with_behavior(TransactionBehavior::Immediate)?;
        let count: i64 = transaction.query_row(
            "SELECT count(*) FROM durable_bindings
             WHERE subject_hash = ?1 AND (revoked_at IS NULL OR revoked_at = ?2)",
            params![binding.subject.as_slice(), PENDING_BINDING],
            |row| row.get(0),
        )?;
        if count >= MAX_BINDINGS_PER_SUBJECT {
            return Err(BindingStoreError::BindingLimit);
        }
        transaction.execute(
            "INSERT INTO durable_bindings (
                binding_id, subject_hash, public_key, host, origin,
                capabilities_json, created_at, last_used_at, revoked_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?7, ?8)",
            params![
                binding.id,
                binding.subject.as_slice(),
                binding.public_key.as_slice(),
                binding.host,
                binding.origin,
                capabilities,
                now,
                revoked_at,
            ],
        )?;
        transaction.commit()?;
        Ok(())
    }

    pub fn activate_pending(
        &self,
        id: &str,
        subject: &[u8; 32],
    ) -> Result<bool, BindingStoreError> {
        let changed = self
            .connection
            .as_ref()
            .ok_or(BindingStoreError::CorruptRecord)?
            .lock()
            .expect("binding database mutex")
            .execute(
                "UPDATE durable_bindings SET revoked_at = NULL, last_used_at = ?3
                 WHERE binding_id = ?1 AND subject_hash = ?2 AND revoked_at = ?4",
                params![id, subject.as_slice(), unix_time()?, PENDING_BINDING],
            )?;
        Ok(changed == 1)
    }

    pub fn get_live(&self, id: &str) -> Result<Option<Binding>, BindingStoreError> {
        #[cfg(test)]
        self.get_live_calls
            .fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let connection = self
            .connection
            .as_ref()
            .ok_or(BindingStoreError::CorruptRecord)?
            .lock()
            .expect("binding database mutex");
        let row = connection
            .query_row(
                "SELECT binding_id, subject_hash, public_key, host, origin, capabilities_json
                 FROM durable_bindings WHERE binding_id = ?1 AND revoked_at IS NULL",
                [id],
                |row| {
                    Ok((
                        row.get::<_, String>(0)?,
                        row.get::<_, Vec<u8>>(1)?,
                        row.get::<_, Vec<u8>>(2)?,
                        row.get::<_, String>(3)?,
                        row.get::<_, String>(4)?,
                        row.get::<_, String>(5)?,
                    ))
                },
            )
            .optional()?;
        let Some((id, subject, public_key, host, origin, capabilities)) = row else {
            return Ok(None);
        };
        let subject = subject
            .try_into()
            .map_err(|_| BindingStoreError::CorruptRecord)?;
        let public_key = public_key
            .try_into()
            .map_err(|_| BindingStoreError::CorruptRecord)?;
        let capabilities: Vec<String> =
            serde_json::from_str(&capabilities).map_err(|_| BindingStoreError::CorruptRecord)?;
        Ok(Some(Binding {
            id,
            subject,
            public_key,
            host,
            origin,
            capabilities,
        }))
    }

    pub fn touch(&self, id: &str) -> Result<bool, BindingStoreError> {
        let changed = self
            .connection
            .as_ref()
            .ok_or(BindingStoreError::CorruptRecord)?
            .lock()
            .expect("binding database mutex")
            .execute(
                "UPDATE durable_bindings SET last_used_at = ?2
             WHERE binding_id = ?1 AND revoked_at IS NULL",
                params![id, unix_time()?],
            )?;
        Ok(changed == 1)
    }

    pub fn revoke(&self, id: &str, subject: &[u8; 32]) -> Result<bool, BindingStoreError> {
        #[cfg(test)]
        if self
            .fail_next_revoke
            .swap(false, std::sync::atomic::Ordering::SeqCst)
        {
            return Err(BindingStoreError::CorruptRecord);
        }
        let mut connection = self
            .connection
            .as_ref()
            .ok_or(BindingStoreError::CorruptRecord)?
            .lock()
            .expect("binding database mutex");
        let transaction = connection.transaction_with_behavior(TransactionBehavior::Immediate)?;
        let existing: Option<(Vec<u8>, Option<i64>)> = transaction
            .query_row(
                "SELECT subject_hash, revoked_at FROM durable_bindings WHERE binding_id = ?1",
                [id],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .optional()?;
        let Some((owner, revoked_at)) = existing else {
            return Ok(false);
        };
        if owner.as_slice() != subject.as_slice() {
            return Ok(false);
        }
        if revoked_at == Some(PENDING_BINDING) {
            transaction.execute(
                "DELETE FROM durable_bindings
                 WHERE binding_id = ?1 AND subject_hash = ?2 AND revoked_at = ?3",
                params![id, subject.as_slice(), PENDING_BINDING],
            )?;
        } else if revoked_at.is_none() {
            transaction.execute(
                "UPDATE durable_bindings SET revoked_at = ?2
                 WHERE binding_id = ?1 AND revoked_at IS NULL",
                params![id, unix_time()?],
            )?;
            transaction.execute(
                "DELETE FROM durable_bindings
                 WHERE binding_id IN (
                    SELECT binding_id FROM durable_bindings
                    WHERE subject_hash = ?1 AND revoked_at IS NOT NULL AND revoked_at != ?3
                    ORDER BY revoked_at DESC, rowid DESC
                    LIMIT -1 OFFSET ?2
                 )",
                params![
                    subject.as_slice(),
                    MAX_REVOKED_BINDINGS_PER_SUBJECT,
                    PENDING_BINDING
                ],
            )?;
        }
        transaction.commit()?;
        Ok(true)
    }

    #[cfg(test)]
    pub fn live_count(&self, subject: &[u8; 32]) -> i64 {
        self.connection
            .as_ref()
            .expect("enabled binding store")
            .lock()
            .expect("binding database mutex")
            .query_row(
                "SELECT count(*) FROM durable_bindings
                 WHERE subject_hash = ?1 AND revoked_at IS NULL",
                params![subject.as_slice()],
                |row| row.get(0),
            )
            .unwrap()
    }

    #[cfg(test)]
    pub fn get_live_call_count(&self) -> usize {
        self.get_live_calls
            .load(std::sync::atomic::Ordering::Relaxed)
    }

    #[cfg(test)]
    pub fn fail_next_revoke(&self) {
        self.fail_next_revoke
            .store(true, std::sync::atomic::Ordering::SeqCst);
    }

    #[cfg(test)]
    fn total_count(&self) -> i64 {
        self.connection
            .as_ref()
            .expect("enabled binding store")
            .lock()
            .expect("binding database mutex")
            .query_row("SELECT count(*) FROM durable_bindings", [], |row| {
                row.get(0)
            })
            .unwrap()
    }
}

fn unix_time() -> Result<i64, BindingStoreError> {
    let seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|_| BindingStoreError::CorruptRecord)?
        .as_secs();
    i64::try_from(seconds).map_err(|_| BindingStoreError::CorruptRecord)
}

#[cfg(unix)]
fn precreate_private_file(path: &Path) -> Result<(), std::io::Error> {
    use std::os::unix::fs::{OpenOptionsExt, PermissionsExt};
    std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .mode(0o600)
        .open(path)?;
    std::fs::set_permissions(path, std::fs::Permissions::from_mode(0o600))
}

#[cfg(not(unix))]
fn precreate_private_file(path: &Path) -> Result<(), std::io::Error> {
    std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .map(drop)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn binding(index: u8) -> Binding {
        Binding {
            id: format!("{index:043}"),
            subject: [3; 32],
            public_key: {
                let mut key = [0; 65];
                key[0] = 4;
                key
            },
            host: "Word".to_owned(),
            origin: "https://office.8-216-134-194.sslip.io".to_owned(),
            capabilities: vec!["agent.v1".to_owned()],
        }
    }

    #[test]
    fn limits_live_bindings_to_twelve_per_subject() {
        let store = BindingStore::open(None).unwrap();
        for index in 0..12 {
            store.enroll(&binding(index)).unwrap();
        }
        assert!(matches!(
            store.enroll(&binding(12)),
            Err(BindingStoreError::BindingLimit)
        ));
        assert!(store.revoke(&binding(0).id, &[3; 32]).unwrap());
        store.enroll(&binding(12)).unwrap();
    }

    #[test]
    fn repeated_enrollment_and_revocation_keeps_database_rows_bounded() {
        let store = BindingStore::open(None).unwrap();
        for index in 0..100_u8 {
            let binding = binding(index);
            store.enroll(&binding).unwrap();
            assert!(store.revoke(&binding.id, &binding.subject).unwrap());
        }
        assert_eq!(store.live_count(&[3; 32]), 0);
        assert_eq!(store.total_count(), MAX_REVOKED_BINDINGS_PER_SUBJECT);
    }

    #[test]
    fn revocation_is_idempotent_only_for_the_authenticated_owner() {
        let store = BindingStore::open(None).unwrap();
        let enrolled = binding(1);
        store.enroll(&enrolled).unwrap();
        assert!(store.revoke(&enrolled.id, &enrolled.subject).unwrap());
        assert!(store.revoke(&enrolled.id, &enrolled.subject).unwrap());
        assert!(!store.revoke(&enrolled.id, &[4; 32]).unwrap());
        assert!(!store.revoke(&binding(2).id, &enrolled.subject).unwrap());
    }

    #[test]
    fn pending_binding_is_not_live_until_activation() {
        let store = BindingStore::open(None).unwrap();
        let pending = binding(7);
        store.enroll_pending(&pending).unwrap();
        assert!(store.get_live(&pending.id).unwrap().is_none());
        assert!(
            store
                .activate_pending(&pending.id, &pending.subject)
                .unwrap()
        );
        assert!(store.get_live(&pending.id).unwrap().is_some());
    }

    #[test]
    fn revoking_a_pending_binding_deletes_the_row_and_releases_the_subject_limit() {
        let store = BindingStore::open(None).unwrap();
        for index in 0..12 {
            store.enroll_pending(&binding(index)).unwrap();
        }
        assert_eq!(store.total_count(), 12);

        let revoked = binding(0);
        assert!(store.revoke(&revoked.id, &revoked.subject).unwrap());
        assert_eq!(store.total_count(), 11);
        store.enroll_pending(&binding(12)).unwrap();
        assert_eq!(store.total_count(), 12);
    }

    #[test]
    fn startup_prunes_uncommitted_pending_bindings() {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "wiswork-relay-pending-{}-{unique}.sqlite",
            std::process::id()
        ));
        let pending = binding(8);
        {
            let store = BindingStore::open(Some(&path)).unwrap();
            store.enroll_pending(&pending).unwrap();
            assert_eq!(store.total_count(), 1);
        }
        let reopened = BindingStore::open(Some(&path)).unwrap();
        assert_eq!(reopened.total_count(), 0);
        drop(reopened);
        std::fs::remove_file(path).unwrap();
    }
}
