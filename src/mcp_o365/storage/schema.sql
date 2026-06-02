-- T4 — Esquema do token store (SQLite).
-- Aplicado no arranque com CREATE TABLE IF NOT EXISTS. Tokens guardados cifrados.
-- Isolamento estrito por utilizador: todas as queries de conta filtram por `subject`.

-- Sessão MCP de um utilizador (1 principal do Plano A -> 1 sessão -> N contas O365).
CREATE TABLE IF NOT EXISTS sessions (
    session_id   TEXT PRIMARY KEY,
    subject      TEXT NOT NULL,
    created_at   TEXT NOT NULL,
    status       TEXT NOT NULL DEFAULT 'active'   -- active | expired
);
CREATE INDEX IF NOT EXISTS idx_sessions_subject ON sessions(subject);

-- Conta O365 ligada (multi-conta desde o início: 1:N por sessão/subject).
CREATE TABLE IF NOT EXISTS linked_accounts (
    account_id        TEXT PRIMARY KEY,
    session_id        TEXT NOT NULL,
    subject           TEXT NOT NULL,
    tenant_id         TEXT,
    home_account_id   TEXT,
    username          TEXT,
    scopes            TEXT,
    access_token_enc  BLOB,            -- cifrado (Cipher)
    refresh_token_enc BLOB,            -- cifrado (Cipher)
    expires_at        TEXT,            -- ISO 8601 UTC
    is_default        INTEGER NOT NULL DEFAULT 0,
    status            TEXT NOT NULL DEFAULT 'active', -- active | expired
    FOREIGN KEY (session_id) REFERENCES sessions(session_id)
);
CREATE INDEX IF NOT EXISTS idx_accounts_subject ON linked_accounts(subject);

-- Clientes OAuth do Plano A registados dinamicamente (RFC 7591, via SDK).
CREATE TABLE IF NOT EXISTS oauth_clients (
    client_id     TEXT PRIMARY KEY,
    metadata_json TEXT NOT NULL,       -- OAuthClientInformationFull serializado
    created_at    TEXT NOT NULL
);

-- Estado transitório do /authorize -> /callback (mapeia `state` ao contexto do pedido).
CREATE TABLE IF NOT EXISTS auth_transactions (
    state                  TEXT PRIMARY KEY,
    client_id              TEXT NOT NULL,
    client_redirect_uri    TEXT NOT NULL,
    client_code_challenge  TEXT,
    client_state           TEXT,
    scopes                 TEXT,
    created_at             TEXT NOT NULL
);

-- Authorization codes do Plano A emitidos ao Claude (curta duração).
CREATE TABLE IF NOT EXISTS authorization_codes (
    code                  TEXT PRIMARY KEY,
    client_id             TEXT NOT NULL,
    subject               TEXT NOT NULL,
    code_challenge        TEXT,
    redirect_uri          TEXT NOT NULL,
    redirect_uri_explicit INTEGER NOT NULL DEFAULT 1,
    scopes                TEXT,
    expires_at            TEXT NOT NULL
);

-- Access tokens do Plano A (opacos) emitidos ao Claude, resolúveis para um subject.
CREATE TABLE IF NOT EXISTS access_tokens (
    token      TEXT PRIMARY KEY,
    client_id  TEXT NOT NULL,
    subject    TEXT NOT NULL,
    scopes     TEXT,
    expires_at TEXT
);

-- Refresh tokens do Plano A (opacos).
CREATE TABLE IF NOT EXISTS refresh_tokens (
    token      TEXT PRIMARY KEY,
    client_id  TEXT NOT NULL,
    subject    TEXT NOT NULL,
    scopes     TEXT,
    expires_at TEXT
);

-- Operações de escrita pendentes do fluxo de aprovação em duas fases (prepare->confirm).
-- O `token` é simultaneamente token de uso único e idempotency key. Payload e resultado
-- são cifrados (Cipher). Isolamento por `subject` em todas as queries.
CREATE TABLE IF NOT EXISTS pending_operations (
    token        TEXT PRIMARY KEY,
    subject      TEXT NOT NULL,
    account_id   TEXT,
    operation    TEXT NOT NULL,
    payload_enc  BLOB NOT NULL,     -- JSON cifrado (Cipher)
    summary      TEXT NOT NULL,
    created_at   TEXT NOT NULL,
    expires_at   TEXT NOT NULL,
    consumed_at  TEXT,              -- NULL até confirmar
    result_enc   BLOB               -- JSON cifrado do resultado (idempotência)
);
CREATE INDEX IF NOT EXISTS idx_pending_subject ON pending_operations(subject);

-- Fase Aprendizagem (US-L.x) — comportamento de email aprendido (só metadados).
-- A assinatura de features é cifrada (Cipher); action/sender_domain/destination ficam em
-- claro para indexação leve (metadados de baixo risco). Isolamento por `subject`.
CREATE TABLE IF NOT EXISTS behavior_events (
    event_id      TEXT PRIMARY KEY,
    subject       TEXT NOT NULL,
    action        TEXT NOT NULL,        -- move|archive|reply|reply_all|forward|delete|send
    sender_domain TEXT,                 -- domínio (sem mailbox local)
    destination   TEXT,                 -- pasta destino quando aplicável
    features_enc  BLOB NOT NULL,        -- JSON cifrado da assinatura (Cipher)
    created_at    TEXT NOT NULL         -- ISO 8601 UTC (retenção/purga)
);
CREATE INDEX IF NOT EXISTS idx_behavior_subject ON behavior_events(subject);
CREATE INDEX IF NOT EXISTS idx_behavior_subject_domain
    ON behavior_events(subject, sender_domain);

-- Consentimento (opt-in) da aprendizagem, por subject. Default: desligado.
CREATE TABLE IF NOT EXISTS learning_preferences (
    subject     TEXT PRIMARY KEY,
    opt_in      INTEGER NOT NULL DEFAULT 0,   -- 0 desligado (default), 1 consentido
    updated_at  TEXT NOT NULL
);

-- Supressões explícitas (feedback "não voltar a sugerir isto"): por (subject, domínio, ação).
-- O recommender ignora padrões aqui listados. Isolamento por `subject`.
CREATE TABLE IF NOT EXISTS learning_suppressions (
    subject       TEXT NOT NULL,
    sender_domain TEXT,                 -- domínio do remetente (NULL = qualquer)
    action        TEXT NOT NULL,        -- ação suprimida (move|archive|reply|...)
    created_at    TEXT NOT NULL,
    PRIMARY KEY (subject, sender_domain, action)
);
