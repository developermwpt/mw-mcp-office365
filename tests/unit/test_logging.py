"""Unit — logging (T2): evento refresh_failure com schema correto e sem PII."""

from __future__ import annotations

import json
import logging

from mcp_o365.logging_setup import (
    JsonFormatter,
    log_refresh_failure,
    set_request_id,
    subject_hash,
)


def _capture(func, logger_name="t"):
    logger = logging.getLogger(logger_name)
    logger.handlers.clear()
    logger.propagate = False
    records: list[str] = []

    class _H(logging.Handler):
        def emit(self, record):
            records.append(JsonFormatter().format(record))

    logger.addHandler(_H())
    logger.setLevel(logging.DEBUG)
    func(logger)
    return [json.loads(r) for r in records]


def test_refresh_failure_schema_sem_pii():
    out = _capture(
        lambda lg: log_refresh_failure(
            lg, subject="user@example.com", account_id="acc-1", reason="invalid_grant"
        )
    )
    assert len(out) == 1
    ev = out[0]
    assert ev["event"] == "refresh_failure"
    assert ev["reason"] == "invalid_grant"
    assert ev["account_id"] == "acc-1"
    # subject pseudonimizado: não aparece o UPN em claro.
    assert "user@example.com" not in json.dumps(ev)
    assert ev["subject_hash"] == subject_hash("user@example.com")


def test_request_id_propaga():
    set_request_id("req-123")
    out = _capture(lambda lg: lg.info("ola"))
    set_request_id(None)
    assert out[0]["request_id"] == "req-123"


def test_subject_hash_estavel_e_truncado():
    h = subject_hash("abc")
    assert len(h) == 16
    assert h == subject_hash("abc")
    assert h != subject_hash("abd")
