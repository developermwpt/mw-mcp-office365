"""Unit — log_audit (v1.1 §1.2): regista metadados (event=audit) sem PII.

Captura-se o registo via caplog e inspecionam-se os `fields` estruturados.
"""

from __future__ import annotations

import logging

from mcp_o365.logging_setup import subject_hash
from mcp_o365.observability.audit import log_audit


def _fields(record: logging.LogRecord) -> dict:
    return getattr(record, "fields", {})


def test_log_audit_emite_evento_com_subject_hash(caplog):
    logger = logging.getLogger("mcp_o365.audit")
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        log_audit(
            logger, action="email.send", subject="user@example.com",
            account_id="acc-1", outcome="success", recipients_count=3,
        )
    records = [r for r in caplog.records if _fields(r).get("event") == "audit"]
    assert len(records) == 1
    f = _fields(records[0])
    assert f["action"] == "email.send"
    assert f["outcome"] == "success"
    assert f["recipients_count"] == 3
    # Pseudonimização: o subject aparece em hash, NUNCA em claro.
    assert f["subject_hash"] == subject_hash("user@example.com")
    assert f["subject_hash"] != "user@example.com"


def test_log_audit_nao_contem_pii_de_enderecos(caplog):
    logger = logging.getLogger("mcp_o365.audit")
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        log_audit(
            logger, action="email.send", subject="user@example.com",
            account_id="acc-1", outcome="success", recipients_count=2,
            extra={"large_attachments": False},
        )
    record = next(r for r in caplog.records if _fields(r).get("event") == "audit")
    blob = repr(_fields(record))
    # Nenhum endereço de email em claro deve aparecer no registo.
    assert "user@example.com" not in blob
    assert "@" not in blob


def test_log_audit_inclui_extra_e_target(caplog):
    logger = logging.getLogger("mcp_o365.audit")
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        log_audit(
            logger, action="email.delete", subject="subj-1",
            account_id="acc-1", target="msg-123", outcome="success",
            extra={"permanent": True},
        )
    record = next(r for r in caplog.records if _fields(r).get("event") == "audit")
    f = _fields(record)
    assert f["target"] == "msg-123"
    assert f["permanent"] is True
    assert "recipients_count" not in f  # não passado -> não aparece
