"""T4 — Token store SQLite, cifrado e isolado por utilizador.

Persiste sessões, contas O365 ligadas (multi-conta), clientes OAuth do Plano A, e o
estado transitório/códigos/tokens do fluxo OAuth. Os tokens Graph são cifrados pelo
`Cipher` antes de tocar o disco. Todas as queries de conta filtram por `subject` — não há
acesso cruzado entre utilizadores.

O relógio é injetado (`clock`) para testabilidade. A ligação SQLite é partilhada e
protegida por um lock (suporta `:memory:` nos testes).
"""

from __future__ import annotations

import json
import sqlite3
import threading
import uuid
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path

from .crypto import Cipher

_SCHEMA = (Path(__file__).parent / "schema.sql").read_text(encoding="utf-8")


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _iso(dt: datetime | None) -> str | None:
    return dt.isoformat() if dt else None


def _parse(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None


class TokenStore:
    """CRUD sobre SQLite com tokens cifrados e isolamento por `subject`."""

    def __init__(
        self,
        path: str,
        cipher: Cipher,
        clock: Callable[[], datetime] = _utcnow,
    ) -> None:
        self._cipher = cipher
        self._clock = clock
        self._lock = threading.Lock()
        self._conn = sqlite3.connect(path, check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA foreign_keys = ON;")
        self._conn.executescript(_SCHEMA)
        self._conn.commit()

    # --- liveness/readiness ---
    def ping(self) -> bool:
        with self._lock:
            self._conn.execute("SELECT 1;")
        return True

    def close(self) -> None:
        self._conn.close()

    # --- sessões ---
    def create_session(self, subject: str) -> str:
        session_id = uuid.uuid4().hex
        with self._lock:
            self._conn.execute(
                "INSERT INTO sessions(session_id, subject, created_at, status) "
                "VALUES (?,?,?, 'active')",
                (session_id, subject, _iso(self._clock())),
            )
            self._conn.commit()
        return session_id

    def get_active_session(self, subject: str) -> sqlite3.Row | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM sessions WHERE subject=? AND status='active' "
                "ORDER BY created_at DESC LIMIT 1",
                (subject,),
            )
            return cur.fetchone()

    def mark_session_expired(self, subject: str) -> None:
        with self._lock:
            self._conn.execute(
                "UPDATE sessions SET status='expired' WHERE subject=?", (subject,)
            )
            self._conn.execute(
                "UPDATE linked_accounts SET status='expired' WHERE subject=?", (subject,)
            )
            self._conn.commit()

    # --- contas ligadas (multi-conta) ---
    def upsert_account(
        self,
        *,
        subject: str,
        session_id: str,
        access_token: str | None,
        refresh_token: str | None,
        expires_at: datetime | None,
        tenant_id: str | None = None,
        home_account_id: str | None = None,
        username: str | None = None,
        scopes: list[str] | None = None,
        is_default: bool = False,
    ) -> str:
        account_id = home_account_id or uuid.uuid4().hex
        at_enc = self._cipher.encrypt_str(access_token) if access_token else None
        rt_enc = self._cipher.encrypt_str(refresh_token) if refresh_token else None
        with self._lock:
            if is_default:
                self._conn.execute(
                    "UPDATE linked_accounts SET is_default=0 WHERE subject=?", (subject,)
                )
            self._conn.execute(
                """
                INSERT INTO linked_accounts(
                    account_id, session_id, subject, tenant_id, home_account_id,
                    username, scopes, access_token_enc, refresh_token_enc, expires_at,
                    is_default, status)
                VALUES (?,?,?,?,?,?,?,?,?,?,?, 'active')
                ON CONFLICT(account_id) DO UPDATE SET
                    session_id=excluded.session_id,
                    tenant_id=excluded.tenant_id,
                    username=excluded.username,
                    scopes=excluded.scopes,
                    access_token_enc=excluded.access_token_enc,
                    refresh_token_enc=excluded.refresh_token_enc,
                    expires_at=excluded.expires_at,
                    is_default=excluded.is_default,
                    status='active'
                """,
                (
                    account_id, session_id, subject, tenant_id, home_account_id,
                    username, " ".join(scopes or []), at_enc, rt_enc, _iso(expires_at),
                    1 if is_default else 0,
                ),
            )
            self._conn.commit()
        return account_id

    def update_account_tokens(
        self,
        *,
        subject: str,
        account_id: str,
        access_token: str,
        refresh_token: str | None,
        expires_at: datetime | None,
    ) -> None:
        at_enc = self._cipher.encrypt_str(access_token)
        rt_enc = self._cipher.encrypt_str(refresh_token) if refresh_token else None
        with self._lock:
            # Filtra por subject E account_id — isolamento.
            self._conn.execute(
                """UPDATE linked_accounts
                   SET access_token_enc=?,
                       refresh_token_enc=COALESCE(?, refresh_token_enc),
                       expires_at=?, status='active'
                   WHERE subject=? AND account_id=?""",
                (at_enc, rt_enc, _iso(expires_at), subject, account_id),
            )
            self._conn.commit()

    def list_accounts(self, subject: str) -> list[dict]:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM linked_accounts WHERE subject=? AND status='active'",
                (subject,),
            )
            rows = cur.fetchall()
        return [self._account_to_dict(r) for r in rows]

    def get_account(self, subject: str, account_id: str) -> dict | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM linked_accounts WHERE subject=? AND account_id=? "
                "AND status='active'",
                (subject, account_id),
            )
            row = cur.fetchone()
        return self._account_to_dict(row) if row else None

    def get_default_account(self, subject: str) -> dict | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM linked_accounts WHERE subject=? AND status='active' "
                "ORDER BY is_default DESC LIMIT 1",
                (subject,),
            )
            row = cur.fetchone()
        return self._account_to_dict(row) if row else None

    def mark_account_expired(self, subject: str, account_id: str) -> None:
        with self._lock:
            self._conn.execute(
                "UPDATE linked_accounts SET status='expired' "
                "WHERE subject=? AND account_id=?",
                (subject, account_id),
            )
            self._conn.commit()

    def _account_to_dict(self, row: sqlite3.Row) -> dict:
        access = (
            self._cipher.decrypt_str(row["access_token_enc"])
            if row["access_token_enc"] else None
        )
        refresh = (
            self._cipher.decrypt_str(row["refresh_token_enc"])
            if row["refresh_token_enc"] else None
        )
        return {
            "account_id": row["account_id"],
            "session_id": row["session_id"],
            "subject": row["subject"],
            "tenant_id": row["tenant_id"],
            "home_account_id": row["home_account_id"],
            "username": row["username"],
            "scopes": (row["scopes"] or "").split(),
            "access_token": access,
            "refresh_token": refresh,
            "expires_at": _parse(row["expires_at"]),
            "is_default": bool(row["is_default"]),
            "status": row["status"],
        }

    # --- clientes OAuth do Plano A (DCR) ---
    def save_client(self, client_id: str, metadata_json: str) -> None:
        with self._lock:
            self._conn.execute(
                "INSERT OR REPLACE INTO oauth_clients(client_id, metadata_json, created_at) "
                "VALUES (?,?,?)",
                (client_id, metadata_json, _iso(self._clock())),
            )
            self._conn.commit()

    def get_client(self, client_id: str) -> str | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT metadata_json FROM oauth_clients WHERE client_id=?", (client_id,)
            )
            row = cur.fetchone()
        return row["metadata_json"] if row else None

    # --- transações de autorização (state) ---
    def save_transaction(
        self,
        *,
        state: str,
        client_id: str,
        client_redirect_uri: str,
        client_code_challenge: str | None,
        client_state: str | None,
        scopes: list[str] | None,
    ) -> None:
        with self._lock:
            self._conn.execute(
                """INSERT OR REPLACE INTO auth_transactions(
                       state, client_id, client_redirect_uri, client_code_challenge,
                       client_state, scopes, created_at)
                   VALUES (?,?,?,?,?,?,?)""",
                (
                    state, client_id, client_redirect_uri, client_code_challenge,
                    client_state, " ".join(scopes or []), _iso(self._clock()),
                ),
            )
            self._conn.commit()

    def pop_transaction(self, state: str) -> dict | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM auth_transactions WHERE state=?", (state,)
            )
            row = cur.fetchone()
            if row:
                self._conn.execute("DELETE FROM auth_transactions WHERE state=?", (state,))
                self._conn.commit()
        if not row:
            return None
        return {
            "state": row["state"],
            "client_id": row["client_id"],
            "client_redirect_uri": row["client_redirect_uri"],
            "client_code_challenge": row["client_code_challenge"],
            "client_state": row["client_state"],
            "scopes": (row["scopes"] or "").split(),
        }

    # --- authorization codes do Plano A ---
    def save_auth_code(
        self,
        *,
        code: str,
        client_id: str,
        subject: str,
        code_challenge: str | None,
        redirect_uri: str,
        redirect_uri_explicit: bool,
        scopes: list[str] | None,
        expires_at: datetime,
    ) -> None:
        with self._lock:
            self._conn.execute(
                """INSERT INTO authorization_codes(
                       code, client_id, subject, code_challenge, redirect_uri,
                       redirect_uri_explicit, scopes, expires_at)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    code, client_id, subject, code_challenge, redirect_uri,
                    1 if redirect_uri_explicit else 0, " ".join(scopes or []),
                    _iso(expires_at),
                ),
            )
            self._conn.commit()

    def get_auth_code(self, code: str) -> dict | None:
        """Lê o authorization code sem o consumir (o SDK faz load antes de exchange)."""
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM authorization_codes WHERE code=?", (code,)
            )
            row = cur.fetchone()
        if not row:
            return None
        return {
            "code": row["code"],
            "client_id": row["client_id"],
            "subject": row["subject"],
            "code_challenge": row["code_challenge"],
            "redirect_uri": row["redirect_uri"],
            "redirect_uri_explicit": bool(row["redirect_uri_explicit"]),
            "scopes": (row["scopes"] or "").split(),
            "expires_at": _parse(row["expires_at"]),
        }

    def delete_auth_code(self, code: str) -> None:
        with self._lock:
            self._conn.execute("DELETE FROM authorization_codes WHERE code=?", (code,))
            self._conn.commit()

    def pop_auth_code(self, code: str) -> dict | None:
        with self._lock:
            cur = self._conn.execute(
                "SELECT * FROM authorization_codes WHERE code=?", (code,)
            )
            row = cur.fetchone()
            if row:
                self._conn.execute("DELETE FROM authorization_codes WHERE code=?", (code,))
                self._conn.commit()
        if not row:
            return None
        return {
            "code": row["code"],
            "client_id": row["client_id"],
            "subject": row["subject"],
            "code_challenge": row["code_challenge"],
            "redirect_uri": row["redirect_uri"],
            "redirect_uri_explicit": bool(row["redirect_uri_explicit"]),
            "scopes": (row["scopes"] or "").split(),
            "expires_at": _parse(row["expires_at"]),
        }

    # --- access / refresh tokens do Plano A ---
    def save_access_token(
        self, *, token: str, client_id: str, subject: str,
        scopes: list[str] | None, expires_at: datetime | None,
    ) -> None:
        with self._lock:
            self._conn.execute(
                "INSERT OR REPLACE INTO access_tokens(token, client_id, subject, scopes, "
                "expires_at) VALUES (?,?,?,?,?)",
                (token, client_id, subject, " ".join(scopes or []), _iso(expires_at)),
            )
            self._conn.commit()

    def get_access_token(self, token: str) -> dict | None:
        return self._get_token("access_tokens", token)

    def delete_access_token(self, token: str) -> None:
        self._delete_token("access_tokens", token)

    def save_refresh_token(
        self, *, token: str, client_id: str, subject: str,
        scopes: list[str] | None, expires_at: datetime | None,
    ) -> None:
        with self._lock:
            self._conn.execute(
                "INSERT OR REPLACE INTO refresh_tokens(token, client_id, subject, scopes, "
                "expires_at) VALUES (?,?,?,?,?)",
                (token, client_id, subject, " ".join(scopes or []), _iso(expires_at)),
            )
            self._conn.commit()

    def get_refresh_token(self, token: str) -> dict | None:
        return self._get_token("refresh_tokens", token)

    def delete_refresh_token(self, token: str) -> None:
        self._delete_token("refresh_tokens", token)

    def _get_token(self, table: str, token: str) -> dict | None:
        with self._lock:
            cur = self._conn.execute(
                f"SELECT * FROM {table} WHERE token=?", (token,)  # noqa: S608 (nome fixo)
            )
            row = cur.fetchone()
        if not row:
            return None
        return {
            "token": row["token"],
            "client_id": row["client_id"],
            "subject": row["subject"],
            "scopes": (row["scopes"] or "").split(),
            "expires_at": _parse(row["expires_at"]),
        }

    def _delete_token(self, table: str, token: str) -> None:
        with self._lock:
            self._conn.execute(
                f"DELETE FROM {table} WHERE token=?", (token,)  # noqa: S608 (nome fixo)
            )
            self._conn.commit()
