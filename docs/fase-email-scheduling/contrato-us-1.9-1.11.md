# Contrato de Implementação — US-1.9 / US-1.10 / US-1.11 (agendamento de envio de email)

**Projeto:** mw-mcp-office365 · **Fase:** Email (extensão do Módulo Email, Fase 1) · **User stories:** US-1.9 (agendar), US-1.10 (listar), US-1.11 (cancelar)
**Documento:** Contrato técnico para o Dev/QA. Decisões de produto fechadas pelo cliente (não reabrir — ver análise §2) e decisões de implementação fechadas pelo PM neste contrato (§0).
**Estado:** Pronto para implementação. Análise funcional fechada em [analise-funcional-agendamento.md](analise-funcional-agendamento.md) (mecanismo `PidTagDeferredSendTime` = `"SystemTime 0x3FEF"` UTC ISO 8601; fluxo draft→send; fonte de verdade = mailbox) e [estado-user-stories.md](estado-user-stories.md).
**Precedências:** US-1.3/US-1.6 (envio + anexos grandes draft→send) já implementadas; US-1.8 (soft/hard delete) já implementada; Fase 2 (`get_mailbox_timezone`, `_resolve_tz` best-effort). Este contrato **reutiliza** integralmente esses caminhos e **não** os altera.

> **Âmbito.** Acrescenta **cinco** tools (`email_schedule_prepare`/`_confirm`, `email_list_scheduled`, `email_schedule_cancel_prepare`/`_confirm`), uma constante de extended property partilhada, **um** método novo no `GraphClient` (`list_deferred_drafts`), e adapta `create_draft`/`get_message` (sem novo endpoint). O Dev segue este contrato à risca. **Não** se escreve código de produção aqui (este é doc; pseudocódigo é ilustrativo). Testes com Graph mockado (FakeGraphClient a estender — §8).

---

## 0. Decisões de implementação fechadas pelo PM (lei deste contrato)

Estas resolvem as ambiguidades deixadas em aberto pela análise. **Não reabrir** sem voltar ao PM.

| # | Decisão | Valor fixado | Justificação |
|---|---------|--------------|--------------|
| **P1** | **Margem mínima** do `send_at` no futuro | **120 segundos (2 minutos)** | Evita a corrida relógio/latência e o "envio imediato" do Exchange quando `send_at` é quase-agora (análise §6.2). Valor fechado: `_MIN_SCHEDULE_MARGIN = timedelta(minutes=2)`. |
| **P2** | **Limite superior** do `send_at` | **365 dias (1 ano)** | Rascunhos diferidos a anos são quase sempre erro de cálculo do assistente. Valor fechado: `_MAX_SCHEDULE_HORIZON = timedelta(days=365)`. |
| **P3** | **Fuso Windows↔IANA / conversão para UTC** | **O servidor NÃO converte fuso. O `send_at` chega já como instante absoluto (com offset ou `Z`); o servidor exige-o e valida-o.** O fuso do mailbox é lido **só para apresentação** no resumo (best-effort) e **nunca** usado para interpretar uma hora "nua". | Ver §5 — análise crítica. A propriedade `SystemTime` é UTC cru (não há conversão server-side do Graph como no `calendarView`); converter uma wall-clock em Python obrigaria a um mapa Windows→IANA frágil. Espelha o princípio "tools recebem valores resolvidos" (calendário/Teams §6.3 da análise). |
| **P4** | **Soft vs hard delete no cancelamento** | **Só soft delete** (mover o rascunho para `deleteditems`). **Sem** `confirm_permanent`/reforço. | É um rascunho do próprio, ainda não enviado; recuperável é o comportamento seguro e suficiente (análise §7.3). Acrescentar reforço seria fricção sem ganho. (Difere do US-1.8, que tinha conteúdo "real" a eliminar.) |
| **P5** | **Evento de aprendizagem no agendar** | **Regista `record_action_event(action="schedule", ...)`** no `confirm` de US-1.9. **Não** regista no cancelar nem no listar. | Coerente com o `send` (que regista `action="send"`). Agendar é uma intenção de envio distinta; usar `action="schedule"` (não reutilizar `"send"`) mantém a aprendizagem honesta. O cancelar é uma correção, não um padrão a aprender. |
| **P6** | **Helper de resolução de fuso** | **Reutilizar** `calendar._resolve_tz` importando-o em `email.py` (NÃO duplicar, NÃO extrair já para um módulo partilhado). | `_resolve_tz` já é genérico, best-effort e degradante (não vive em estado de calendário). `email.py` já importa `_domains` de... na verdade é `calendar.py` que importa de `email.py`; aqui inverte-se. Para evitar import circular (email↔calendar), ver §5.3: extrair `_resolve_tz` para `tools/_timezone.py`. **Decisão: extrair para `tools/_timezone.py`** e fazer ambos (`calendar.py`, `email.py`) importarem de lá. |
| **P7** | **`get_message` traz a extended property no `cancel_prepare`** | **Sim, best-effort.** `get_message` ganha um parâmetro opcional `expand` para pedir `singleValueExtendedProperties`; o `cancel_prepare` usa-o para validar que o rascunho ainda é diferido e montar o resumo. Falha de leitura → degradação (resumo genérico, ainda emite token). | Permite o resumo informativo (hora/assunto) e o `error` orientador quando já não é diferido (análise US-1.11 AC3), sem novo endpoint. |

---

## 1. Núcleo — `src/mcp_o365/tools/email.py`

### 1.0 Imports e constantes a acrescentar (topo do módulo)

```python
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError  # só para apresentação (P3/§5)
from ._timezone import _resolve_tz                      # extraído de calendar.py (P6/§5.3)

# Extended property MAPI PidTagDeferredSendTime (envio diferido nativo do Exchange).
# id = "{graph_type} {proptag}"; PT_SYSTIME -> SystemTime; tag 0x3FEF. Valor SEMPRE UTC ISO 8601.
_DEFERRED_SEND_PROP_ID = "SystemTime 0x3FEF"

# Validação temporal do agendamento (P1/P2).
_MIN_SCHEDULE_MARGIN = timedelta(minutes=2)    # >= 2 min no futuro
_MAX_SCHEDULE_HORIZON = timedelta(days=365)    # <= 1 ano
```

> **Constante partilhada (P3 do briefing):** `_DEFERRED_SEND_PROP_ID` vive em `email.py` (é o módulo dono do agendamento). O `GraphClient` **não** importa de `tools/` (camada inferior) — por isso a query da listagem (§3.2) recebe o `prop_id` como argumento vindo da tool, ou redefine a string literal no `client.py` com um comentário a apontar para a constante. **Decisão: a tool passa `prop_id=_DEFERRED_SEND_PROP_ID` ao método do client** (mantém a camada Graph agnóstica).

### 1.1 `run_email_schedule_prepare` — assinatura exata

Baseada em `run_email_send_prepare` (linhas ~397-455), **mais** `send_at` (obrigatório) e `timezone` (opcional, só apresentação). Mantém a ordem dos parâmetros existentes; os novos entram **keyword-only** depois de `body_type`/`attachments` e antes de `message_meta` (todos já são keyword-only após o `*`, logo a ordem relativa é livre — colocá-los a seguir a `attachments` para legibilidade):

```python
async def run_email_schedule_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,           # <-- NOVO face ao send_prepare (precisa para _resolve_tz)
    store: TokenStore,
    approval: ApprovalEngine,
    to: list[str],
    body: str,
    send_at: str,                        # <-- NOVO: instante ABSOLUTO ISO 8601 (offset/Z). Obrigatório.
    subject_line: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_type: str = "Text",
    attachments: list[dict] | None = None,
    timezone: str | None = None,         # <-- NOVO: fuso explícito do utilizador (só apresentação)
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

> **Nota:** `run_email_send_prepare` **não** recebe `graph_client` (só resolve token); o `schedule_prepare` recebe-o porque resolve o fuso do mailbox via `_resolve_tz` (best-effort). É o único acréscimo de dependência.

**Fluxo (pseudocódigo fiel ao estilo do `email.py`):**

```python
    """US-1.9 — Prepara o agendamento: valida destinatários e a HORA, monta a mensagem com a
    extended property de envio diferido, devolve token. NÃO escreve no Graph."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if not to:
        return {"status": "error", "message": "É obrigatório indicar destinatários (to)."}

    # --- Validação temporal (P1/P2, análise §6.2) — ANTES de qualquer escrita, sem token. ---
    when = _parse_iso(send_at)                       # reutiliza o helper existente (aceita 'Z')
    if when is None:
        return {"status": "error", "message":
                "send_at inválido: indique um instante ISO 8601 com fuso/offset (ex.: "
                "2026-06-10T09:00:00+01:00 ou ...Z)."}
    # P3: exigir instante ABSOLUTO. _parse_iso já completa naive->UTC, por isso validamos
    # explicitamente a presença de offset no texto (ver helper _has_offset em §1.5).
    if not _has_offset(send_at):
        return {"status": "error", "message":
                "send_at tem de incluir o fuso/offset resolvido (ex.: +01:00 ou Z). "
                "Resolva a hora no fuso do mailbox a montante."}
    now = clock()
    delta = when - now
    if delta < _MIN_SCHEDULE_MARGIN:
        return {"status": "error", "message":
                "A hora de envio tem de estar pelo menos 2 minutos no futuro."}
    if delta > _MAX_SCHEDULE_HORIZON:
        return {"status": "error", "message":
                "A hora de envio não pode estar a mais de 1 ano de distância."}

    send_at_utc = _to_utc_iso(when)                  # normaliza para "...Z" (ver §1.5)

    # --- Fuso só para APRESENTAÇÃO (best-effort; P3/§5). Nunca interpreta a hora. ---
    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )  # pode ser None (fuso indisponível) -> apresentar em UTC e declará-lo.
    when_label = _present_in_tz(when, tz_label)      # string humana (ver §1.5)

    # --- Montagem da message + extended property (mesmo _build_message do send). ---
    message = _build_message(
        to=to, cc=cc, bcc=bcc, subject=subject_line, body=body,
        body_type=body_type, attachments=attachments,
    )
    message["singleValueExtendedProperties"] = [
        {"id": _DEFERRED_SEND_PROP_ID, "value": send_at_utc}
    ]

    total = len(to) + len(cc or []) + len(bcc or [])
    large = _attachment_too_large(attachments)
    summary = (
        f"Agendar email para {total} destinatário(s) "
        f"(domínios: {', '.join(_domains(to)) or 'n/d'}), "
        f"assunto '{subject_line or '(sem assunto)'}', envio em {when_label}."
    )
    if large:
        summary += " Inclui anexo(s) grande(s) (envio via upload session)."
    if tz_label is None and timezone is None:
        summary += " (Fuso do mailbox indisponível; hora interpretada/apresentada em UTC.)"

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.schedule",                  # <-- operação NOVA
        payload={
            "message": message,                      # já inclui singleValueExtendedProperties
            "recipients_count": total,
            "large_attachments": large,
            "send_at_utc": send_at_utc,              # para auditoria (só-metadados)
            "message_meta": _safe_meta(message_meta),
        },
        summary=summary,
    )
    prepared["recipients_count"] = total
    prepared["large_attachments"] = large
    prepared["send_at_utc"] = send_at_utc
    return prepared
```

**Notas de fidelidade:**
- O `message` no payload **já** transporta `singleValueExtendedProperties` — o `confirm` não precisa de o recompor (o draft é criado com ele).
- A validação temporal corre **antes** de `approval.prepare` → recusa **sem token** (invariante de §10 da análise: prepare não escreve, e validação falhada não emite token).
- `_resolve_tz` nunca passa por `call_graph`; uma falha de fuso **não** derruba a sessão (lição Fase 2) — devolve `None` e o resumo declara UTC.

### 1.2 `run_email_schedule_confirm` — assinatura exata

Idêntica a `run_email_send_confirm` (linhas ~458-526). Não há parâmetros novos.

```python
async def run_email_schedule_confirm(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    confirmation_token: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

**Executor (pseudocódigo) — reutiliza o caminho draft→send do send_confirm, SEMPRE (nunca `send_mail`):**

```python
    """US-1.9 — Confirma o agendamento com token fresco; audita `email.schedule`."""
    async def executor(operation: str, payload: dict) -> dict:
        message = payload["message"]                 # inclui singleValueExtendedProperties

        async def do(token: str) -> str | None:
            # Diferido OBRIGA draft->send (a extended property é definida na CRIAÇÃO do rascunho;
            # `sendMail` ignora-a — análise §3.2). Mesmo sem anexos grandes, usa-se draft->send.
            inline = [a for a in message.get("attachments", []) if not _att_is_large(a)]
            draft = await graph_client.create_draft(token, {**message, "attachments": inline})
            draft_id = draft.get("id")
            if payload.get("large_attachments"):
                for att in message.get("attachments", []):
                    if not _att_is_large(att):
                        continue
                    raw = base64.b64decode(att.get("contentBytes") or "")
                    session = await graph_client.create_attachment_upload_session(
                        token, draft_id,
                        attachment_item={"attachmentType": "file",
                                         "name": att.get("name"), "size": len(raw)},
                    )
                    upload_url = session.get("uploadUrl") if session else None
                    if upload_url:
                        await graph_client.upload_attachment_bytes(upload_url, raw)
            await graph_client.send_draft(token, draft_id)
            return draft_id

        account, draft_id = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=do, account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.schedule", subject=subject,
            account_id=account.account_id, target=draft_id,
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={"large_attachments": bool(payload.get("large_attachments")),
                   "send_at_utc": payload.get("send_at_utc"), "deferred": True},
        )
        record_action_event(
            subject, store=store, action="schedule",      # P5
            message=payload.get("message_meta"), clock=clock,
        )
        return {"operation": operation, "message_id": draft_id,
                "send_at_utc": payload.get("send_at_utc"),
                "message": "Envio agendado."}

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)
```

> **Diferença importante face ao `send_confirm`:** o agendamento usa **sempre** `create_draft`+`send_draft` (mesmo sem anexos grandes), porque a extended property só "pega" no caminho draft→send. **Nunca** chamar `send_mail` aqui. O `confirm` devolve o `message_id` do rascunho diferido (necessário para US-1.11).

### 1.3 `run_email_list_scheduled` — assinatura exata (leitura, sem aprovação)

Estilo de `run_email_search` (leitura via `call_graph`, sanitização, `reauth_response`). **Sem** `approval`.

```python
async def run_email_list_scheduled(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    top: int = 50,
    timezone: str | None = None,         # fuso explícito p/ apresentação (senão mailbox)
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

**Fluxo (pseudocódigo):**

```python
    """US-1.10 — Lista os rascunhos com a extended property de envio diferido cujo instante
    ainda é FUTURO (pendentes). Leitura; não escreve; não exige aprovação."""
    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )
    try:
        _, page = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.list_deferred_drafts(
                token, prop_id=_DEFERRED_SEND_PROP_ID, top=top,
            ),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    now = clock()
    items = []
    for d in page["drafts"]:
        send_at_utc = d.get("deferred_send_at")          # valor da prop expandida (UTC ISO)
        when = _parse_iso(send_at_utc)
        if when is None or when <= now:                  # FILTRO "ainda futuro" CLIENT-SIDE (gate §11)
            continue
        recips = d.get("to") or []
        items.append({
            "id": d.get("id"),
            "subject": sanitize_html(d.get("subject") or ""),  # conteúdo não-confiável
            "recipients_count": len(recips),
            "recipient_domains": _domains(recips),
            "send_at": _present_in_tz(when, tz_label),   # apresentação no fuso
            "send_at_utc": send_at_utc,
        })
    items.sort(key=lambda i: i["send_at_utc"])            # mais próximos primeiro
    return {
        "status": "ok",
        "scheduled": items,
        "count": len(items),
        "has_more": page["next"] is not None,
        "content_is_untrusted": True,
    }
```

**Notas:**
- O corpo **não** é devolvido (minimização). O assunto é sanitizado (`sanitize_html`) e marca-se `content_is_untrusted=true` (defesa em profundidade — análise US-1.10 AC5).
- O **filtro "ainda futuro" é client-side** sobre o valor expandido (análise nota de robustez / gate §11): o `$filter` do Graph só filtra **presença** da propriedade; a comparação por data não é fiável em todos os tenants.
- Paginação consciente: devolve a 1ª página (`top`) + `has_more`. O nº de agendamentos pendentes é tipicamente pequeno; não auto-paginar (mais simples e suficiente — análise US-1.10 AC4).

### 1.4 `run_email_schedule_cancel_prepare` / `_confirm` — assinaturas exatas

Padrão de `run_email_delete_*` (US-1.8), mas **só soft delete** (P4 — sem `permanent`/`confirm_permanent`).

```python
async def run_email_schedule_cancel_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,           # <-- precisa para o get_message best-effort (P7)
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    timezone: str | None = None,         # apresentação
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

**Fluxo (pseudocódigo):**

```python
    """US-1.11 — Prepara o cancelamento de um envio agendado (NÃO elimina). Confirma
    best-effort que o `message_id` é um rascunho ainda diferido e monta o resumo."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    # Best-effort (P7): ler o rascunho com a extended property para validar/apresentar.
    deferred_at = None
    subj = None
    try:
        _, msg = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_message(
                token, message_id, expand=f"singleValueExtendedProperties($filter=id eq '{_DEFERRED_SEND_PROP_ID}')",
            ),
            account_id=account_id, clock=clock,
        )
        deferred_at = _extract_deferred_value(msg)       # None se a prop não estiver presente
        subj = msg.get("subject")
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    # (Outras falhas de leitura: degradar — não bloquear o cancelamento; resumo genérico.)

    # Se conseguimos ler e o rascunho JÁ NÃO é diferido (sem a prop / já enviado) -> erro, sem token.
    if subj is not None and deferred_at is None:
        return {"status": "error", "message":
                "Este email já não é um envio agendado pendente (pode já ter sido enviado "
                "ou cancelado). Liste os agendados com email_list_scheduled."}

    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )
    when = _parse_iso(deferred_at)
    when_label = _present_in_tz(when, tz_label) if when else "hora desconhecida"
    summary = (
        f"Cancelar o envio agendado para {when_label}, "
        f"assunto '{sanitize_html(subj or '(desconhecido)')}'. "
        "O rascunho vai para Itens Eliminados (recuperável)."
    )

    return approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.schedule_cancel",              # <-- operação NOVA
        payload={"message_id": message_id, "message_meta": _safe_meta(message_meta)},
        summary=summary,
    )
```

```python
async def run_email_schedule_cancel_confirm(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    confirmation_token: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

**Executor (pseudocódigo) — soft delete via `move_message` para `deleteditems` (P4):**

```python
    """US-1.11 — Confirma o cancelamento; soft delete do rascunho; audita `email.schedule_cancel`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, moved = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.move_message(
                token, payload["message_id"], destination_id="deleteditems"),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.schedule_cancel", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"permanent": False},
        )
        result = {"operation": operation, "message": "Envio agendado cancelado "
                  "(rascunho movido para Itens Eliminados)."}
        if isinstance(moved, dict) and moved.get("id"):
            result["new_id"] = moved["id"]
        return result

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)
```

> **P4 — sem reforço.** Ao contrário de `run_email_delete_confirm`, **não** há ramo `permanent`/`confirm_permanent`. É sempre soft. (Se um dia se quiser hard, será uma extensão; fora de âmbito.)

### 1.5 Helpers novos em `email.py` (apresentação e validação de offset)

```python
def _has_offset(value: str) -> bool:
    """True se a string ISO traz fuso explícito (Z ou ±HH:MM). P3: send_at TEM de o trazer."""
    v = value.strip()
    return v.endswith("Z") or v.endswith("z") or ("T" in v and ("+" in v[10:] or "-" in v[10:]))

def _to_utc_iso(dt: datetime) -> str:
    """datetime aware -> 'YYYY-MM-DDTHH:MM:SSZ' em UTC (formato aceite pela propriedade)."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def _present_in_tz(dt: datetime, tz_label: str | None) -> str:
    """Apresentação humana da hora no fuso (best-effort). NUNCA interpreta — só formata.
    Aceita IANA ('Europe/Lisbon') via zoneinfo; se for nome Windows ('GMT Standard Time')
    ou inválido, cai para UTC e anota. P3/§5: não é caminho crítico."""
    if tz_label:
        try:
            local = dt.astimezone(ZoneInfo(tz_label))
            return f"{local.strftime('%d/%m/%Y %H:%M')} ({tz_label})"
        except (ZoneInfoNotFoundError, ValueError, KeyError):
            pass  # nome Windows ou inválido -> UTC
    return f"{dt.astimezone(timezone.utc).strftime('%d/%m/%Y %H:%M')} (UTC)"

def _extract_deferred_value(msg: dict) -> str | None:
    """Lê o value da extended property PidTagDeferredSendTime de um get_message expandido."""
    for ep in msg.get("singleValueExtendedProperties") or []:
        if ep.get("id") == _DEFERRED_SEND_PROP_ID:
            return ep.get("value")
    return None
```

> **Sobre `_present_in_tz` e Windows:** é **só apresentação** (P3). Não há mapa Windows→IANA obrigatório. Se o mailbox devolver "GMT Standard Time" (Windows), o `ZoneInfo` falha e cai para UTC com a anotação — aceitável, pois a hora **correta** (instante absoluto) já está garantida pelo `send_at` com offset. **Opcional (não bloqueante):** o Dev pode acrescentar um mini-mapa dos 3-4 fusos mais comuns do cliente (ex.: `{"GMT Standard Time": "Europe/London", "Romance Standard Time": "Europe/Paris"}`) para a apresentação ficar bonita; não é requisito.

---

## 2. Auditoria (regra A1 — só metadados, sem PII, sem 2º subject_hash)

Assinatura **real** (confirmada em `src/mcp_o365/observability/audit.py`):

```python
def log_audit(logger, *, action, subject, account_id=None, target=None,
              outcome, recipients_count=None, extra=None) -> None
```

`extra` é fundido por `fields.update(extra)` **depois** do `subject_hash` de topo. **Regra A1:** NUNCA pôr `subject_hash`, emails, nomes, assunto em claro ou corpo em `extra`.

| Tool (confirm) | `action` | `target` | `recipients_count` | `extra` (só metadados) |
|---|---|---|---|---|
| US-1.9 | `email.schedule` | `draft_id` (id do rascunho diferido) | sim | `{large_attachments: bool, send_at_utc: "...Z", deferred: true}` |
| US-1.11 | `email.schedule_cancel` | `message_id` | — | `{permanent: false}` |

- `send_at_utc` é um instante (não é PII de conteúdo) — permitido em `extra`.
- US-1.10 (listar) é **leitura** → **sem** auditoria de escrita.
- `audit_logger = logging.getLogger("mcp_o365.audit")` já existe no módulo (linha 39).

---

## 3. GraphClient — `src/mcp_o365/graph/client.py`

### 3.1 `create_draft` — **já serve**, sem alteração

`create_draft(self, access_token, message)` faz `POST /me/messages` com o `message` cru (linhas ~300-305). O `message` pode incluir `singleValueExtendedProperties` — o Graph aceita-o no corpo de criação do rascunho. **Nenhuma alteração necessária.** (Confirma o ponto da análise §3.3.)

### 3.2 `list_deferred_drafts` — **método NOVO**

`GET /me/mailFolders/drafts/messages` com `$filter` (presença da prop) + `$expand` (trazer o valor) + `$select` (minimizar).

```python
async def list_deferred_drafts(
    self, access_token: str, *, prop_id: str, top: int = 50
) -> dict:
    """`GET /me/mailFolders/drafts/messages` filtrando rascunhos com a extended property
    `prop_id` (PidTagDeferredSendTime) e expandindo o seu valor (hora de envio diferido).

    O `$filter` testa só a PRESENÇA da propriedade (comparar a data por `ep/value` não é
    fiável em todos os tenants — análise §US-1.10/gate §11); o "ainda futuro" é filtrado
    client-side na tool. Devolve {"drafts": [...], "next": @odata.nextLink}."""
    ep_filter = f"singleValueExtendedProperties/any(ep: ep/id eq '{prop_id}')"
    ep_expand = f"singleValueExtendedProperties($filter=id eq '{prop_id}')"
    params = {
        "$filter": ep_filter,
        "$expand": ep_expand,
        "$select": "id,subject,toRecipients",
        "$top": top,
    }
    data = await self._request(
        "GET", "/me/mailFolders/drafts/messages", access_token, params=params
    ) or {}
    return {
        "drafts": [self._map_deferred_draft(m, prop_id) for m in data.get("value", [])],
        "next": data.get("@odata.nextLink"),
    }

@classmethod
def _map_deferred_draft(cls, m: dict, prop_id: str) -> dict:
    """Mapeia um rascunho diferido: id, subject, destinatários (emails) e o valor da prop."""
    deferred_at = None
    for ep in m.get("singleValueExtendedProperties") or []:
        if ep.get("id") == prop_id:
            deferred_at = ep.get("value")
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "to": [cls._addr(r) for r in m.get("toRecipients", [])],
        "deferred_send_at": deferred_at,     # UTC ISO 8601 (cru, como veio do Exchange)
    }
```

> **Gate §11 (validação manual):** a fiabilidade do `$filter`/`$expand` por extended property na pasta drafts é um gate no tenant. O código está desenhado para degradar (filtra presença; "futuro" client-side), mas **a query tem de ser validada no tenant real** (⬜).

### 3.3 `get_message` — acrescentar parâmetro `expand` (P7)

Atual: `get_message(self, access_token, message_id, *, select=None)` (linhas ~155-175). Acrescentar `expand: str | None = None` e propagá-lo como `$expand`, e **incluir `singleValueExtendedProperties` no dict devolvido** quando presente:

```python
async def get_message(
    self, access_token: str, message_id: str, *, select=None, expand=None
) -> dict:
    params = {}
    if select: params["$select"] = select
    if expand: params["$expand"] = expand
    data = await self._request("GET", f"/me/messages/{message_id}", access_token,
                               params=(params or None)) or {}
    result = { ...campos atuais inalterados... }
    if "singleValueExtendedProperties" in data:           # P7: expor para o cancel_prepare
        result["singleValueExtendedProperties"] = data["singleValueExtendedProperties"]
    return result
```

**Retrocompatível:** sem `expand`, comportamento idêntico ao atual (os chamadores existentes de `get_message` não passam `expand`).

### 3.4 Cancelamento — **reutiliza `move_message`** (sem método novo)

`move_message(..., destination_id="deleteditems")` (linhas ~269-279) já faz o soft delete. **Sem alteração.** (`permanent_delete` **não** é usado — P4.)

---

## 4. `tools/_timezone.py` — extração de `_resolve_tz` (P6)

**Decisão:** extrair `_resolve_tz` (atualmente em `calendar.py` linhas ~74-101) para um módulo partilhado **`src/mcp_o365/tools/_timezone.py`**, e fazer `calendar.py` **e** `email.py` importarem de lá. Razão: `email.py` não pode importar de `calendar.py` sem risco de import circular (o `calendar.py` já importa `_domains` de `email.py`).

- Mover a função tal e qual (mesma assinatura keyword-only, mesma semântica best-effort/degradação para `None`).
- Em `calendar.py`: substituir a definição local por `from ._timezone import _resolve_tz` (mantendo o nome — os call sites em calendar não mudam).
- Em `email.py`: `from ._timezone import _resolve_tz`.
- O `_utcnow` continua em cada módulo (já duplicado hoje; não é objeto deste contrato unificá-lo).

> Se o Dev preferir **não** extrair (para minimizar o diff em calendar.py), a alternativa é replicar `_resolve_tz` em `email.py`. **Não recomendado** (duplicação de lógica best-effort sensível). A extração é a opção fechada.

---

## 5. Fuso horário — análise crítica e decisão (P3)

### 5.1 Porque o calendário NÃO ajuda a resolver isto

`calendar.py` **nunca** converte fuso em Python. Monta `start/end` como wall-clock + `timeZone: tz` (Windows **ou** IANA) e **deixa o Graph converter** (`calendarView` com header `Prefer`, `create_event` com `start.timeZone`). O `get_mailbox_timezone` devolve o valor **tal e qual** ("GMT Standard Time", Windows) e o Graph aceita-o (comentário em `client.py` ~362: "não convertemos").

**O agendamento não tem essa sorte:** a extended property `PidTagDeferredSendTime` é um `SystemTime` **UTC cru** — não há campo `timeZone` ao lado nem header `Prefer` que o Graph interprete. Quem calcula o instante UTC somos **nós**. Converter uma wall-clock no fuso do mailbox em Python obrigaria a mapear "GMT Standard Time" → "Europe/London" (Windows→IANA), porque `zoneinfo` só conhece IANA. Esse mapa é frágil e incompleto.

### 5.2 Decisão (P3): o `send_at` chega absoluto

- **A tool exige `send_at` com offset/`Z`** (instante absoluto). O assistente — que já traduz linguagem natural e conhece o fuso do mailbox/utilizador — resolve a hora **a montante** e envia `2026-06-10T09:00:00+01:00` (ou `...Z`). Espelha o princípio "tools recebem valores resolvidos" do calendário/Teams (análise §6.3).
- **O servidor valida** (offset presente, futuro+margem, limite) e **normaliza para UTC** (`_to_utc_iso`) — conversão trivial e correta porque o instante já é absoluto.
- **O fuso do mailbox é lido só para apresentação** no resumo (`_present_in_tz`, best-effort). Se falhar (Windows ou indisponível), apresenta-se em UTC e **declara-se** no resumo. A hora **gravada** está sempre correta (não depende da apresentação).

Isto resolve a subtileza Windows↔IANA **eliminando-a do caminho crítico**: a correção do instante não depende de nenhum mapa de fusos.

### 5.3 Reforço nas descrições/instructions (§6/§7)

As descrições das tools e as `instructions` do servidor instruem o LLM a **resolver a hora no fuso do mailbox a montante** e a passar `send_at` com offset. É a peça que torna o P3 robusto na prática.

---

## 6. `server.py` — wrappers e descrições das 5 tools novas

Inserir na zona `email_*` (a seguir ao bloco de `email_delete_*`, ~linha 306), seguindo o padrão `@mcp.tool(description=...)` + wrapper assíncrono com `_subject()` e dependências injetadas.

### 6.1 `email_schedule_prepare`

```python
@mcp.tool(
    description=(
        "FASE 1/2 — Prepara o AGENDAMENTO do envio de um email (NÃO agenda nem envia). "
        "send_at TEM de ser um instante ISO 8601 com fuso/offset JÁ RESOLVIDO no fuso do "
        "mailbox do utilizador (ex.: 2026-06-10T09:00:00+01:00 ou ...Z) — resolva a hora a "
        "montante e CONFIRME a hora absoluta com o utilizador antes de chamar. A hora tem de "
        "estar entre 2 minutos e 1 ano no futuro (senão devolve error sem token). O corpo é "
        "conteúdo do utilizador. Devolve resumo + confirmation_token; chame "
        "email_schedule_confirm. O Exchange entrega na hora marcada mesmo com o servidor "
        "desligado."
    )
)
async def email_schedule_prepare(
    to: list[str],
    body: str,
    send_at: str,
    subject_line: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_type: str = "Text",
    attachments: list[dict] | None = None,
    timezone: str | None = None,
    message_meta: dict | None = None,
) -> dict:
    return await email_tools.run_email_schedule_prepare(
        _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
        store=store, approval=approval, to=to, body=body, send_at=send_at,
        subject_line=subject_line, cc=cc, bcc=bcc, body_type=body_type,
        attachments=attachments, timezone=timezone, message_meta=message_meta,
    )
```

### 6.2 `email_schedule_confirm`

```python
@mcp.tool(
    description="FASE 2/2 — Confirma e agenda o envio preparado (requer confirmation_token)."
)
async def email_schedule_confirm(confirmation_token: str) -> dict:
    return await email_tools.run_email_schedule_confirm(
        _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
        store=store, approval=approval, confirmation_token=confirmation_token,
    )
```

### 6.3 `email_list_scheduled`

```python
@mcp.tool(
    description=(
        "Lista os envios de email AGENDADOS e ainda PENDENTES (rascunhos com hora de envio "
        "diferida ainda no futuro). Leitura — não agenda nem cancela. Devolve, por item: id "
        "(use-o em email_schedule_cancel_prepare), assunto (sanitizado), nº de destinatários "
        "e domínios, e a hora de envio no fuso do mailbox + em UTC. O assunto é conteúdo "
        "NÃO-confiável (content_is_untrusted)."
    )
)
async def email_list_scheduled(top: int = 50, timezone: str | None = None) -> dict:
    return await email_tools.run_email_list_scheduled(
        _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
        store=store, top=top, timezone=timezone,
    )
```

### 6.4 `email_schedule_cancel_prepare`

```python
@mcp.tool(
    description=(
        "FASE 1/2 — Prepara o CANCELAMENTO de um envio agendado pendente (NÃO cancela). "
        "message_id é o id do rascunho diferido (de email_list_scheduled ou do retorno de "
        "email_schedule_confirm). O rascunho vai para Itens Eliminados (recuperável). Devolve "
        "resumo + confirmation_token; chame email_schedule_cancel_confirm. NOTA: cancelar "
        "MUITO PERTO da hora de envio pode já não impedir a entrega (o Exchange pode já ter "
        "processado a mensagem)."
    )
)
async def email_schedule_cancel_prepare(
    message_id: str, timezone: str | None = None, message_meta: dict | None = None
) -> dict:
    return await email_tools.run_email_schedule_cancel_prepare(
        _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
        store=store, approval=approval, message_id=message_id,
        timezone=timezone, message_meta=message_meta,
    )
```

### 6.5 `email_schedule_cancel_confirm`

```python
@mcp.tool(
    description="FASE 2/2 — Confirma o cancelamento do envio agendado (requer confirmation_token)."
)
async def email_schedule_cancel_confirm(confirmation_token: str) -> dict:
    return await email_tools.run_email_schedule_cancel_confirm(
        _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
        store=store, approval=approval, confirmation_token=confirmation_token,
    )
```

### 6.6 `instructions` do servidor (recomendado)

Acrescentar uma frase à nota de email das `instructions` (zona ~linha 84+): *"Para AGENDAR um envio (`email_schedule_prepare`), resolva a hora no fuso do mailbox a montante e passe `send_at` em ISO 8601 com offset/Z; confirme a hora absoluta com o utilizador. Para listar/cancelar agendamentos use `email_list_scheduled` e `email_schedule_cancel_*`. Cancelar muito perto da hora pode não impedir o envio."* Não bloqueante (a garantia vem da validação no `prepare`), mas reforça o P3.

---

## 7. Estratégia de teste (QA)

**Ficheiros:** estender `tests/integration/test_email_write_e2e.py` (agendar/cancelar + transversais), `tests/integration/test_email_read_e2e.py` (listar), `tests/unit/test_graph_email_client.py` (montagem do draft + query da listagem + `get_message` com expand), `tests/unit/test_audit.py` (ações novas). Helpers existentes: `_link`, `_approval`, `_plane_b`, `_audit_events`, `gc.count(...)`.

### 7.1 Acréscimos OBRIGATÓRIOS ao `FakeGraphClient` (`tests/integration/fake_graph.py`)

O `FakeGraphClient` já tem `create_draft`, `create_attachment_upload_session`, `upload_attachment_bytes`, `send_draft`, `move_message`, `get_message`, `get_mailbox_timezone`, `me`. **Falta:**

1. **`list_deferred_drafts`** (novo método + atributo programável):
   ```python
   # no __init__: deferred_drafts: dict | None = None
   self._deferred_drafts = deferred_drafts or {"drafts": [], "next": None}
   async def list_deferred_drafts(self, access_token, *, prop_id, top=50) -> dict:
       self._record("list_deferred_drafts", access_token, prop_id=prop_id, top=top)
       return self._deferred_drafts
   ```
   Formato de cada draft: `{"id","subject","to":[...emails], "deferred_send_at":"...Z"}` (já mapeado).
2. **`get_message` aceitar `expand`** (a assinatura atual é `(access_token, message_id, **kwargs)` — já tolera `expand` via `**kwargs`; **confirmar** e fazer o `_message` programável incluir `singleValueExtendedProperties` para os testes de `cancel_prepare`). Não é preciso método novo, só garantir que o `_message` de teste pode trazer a coleção.

   > Nada mais a acrescentar para o draft diferido: o `create_draft` do fake já regista a `message` (com `singleValueExtendedProperties`) e devolve `self._draft` (`{"id":"draft-1"}`) — basta inspecionar `gc.calls` para verificar a extended property e usar `draft-1` como `message_id` retornado.

### 7.2 Casos numerados (mapeados aos critérios de aceitação)

**US-1.9 — Agendar (AC1-AC9):**

1. **T1 (AC1) — prepare não toca no Graph:** `schedule_prepare(send_at=futuro_válido)` → `pending_confirmation` com `confirmation_token`; `gc.calls == []` (nem `create_draft` nem `send_draft` nem `get_mailbox_timezone`... **exceto**: `_resolve_tz` chama `get_mailbox_timezone` — ver T1b). `recipients_count` e `large_attachments` presentes.
   - **T1b:** o `prepare` pode chamar `get_mailbox_timezone` (apresentação). Verificar que `create_draft`/`send_draft` estão a **0**; `get_mailbox_timezone` ≤ 1 é aceitável (leitura best-effort, não-escrita).
2. **T2 (AC6) — confirm executa draft→send com a propriedade:** `confirm(token)` → `done`; `gc.count("create_draft")==1`, `gc.count("send_draft")==1`, `gc.count("send_mail")==0`. Inspecionar `gc.calls`: a `message` do `create_draft` contém `singleValueExtendedProperties == [{"id":"SystemTime 0x3FEF","value":"<...Z>"}]`. Resposta inclui `message_id` (= "draft-1").
3. **T3 (AC1/idempotência) — replay:** segundo `confirm(token)` → `idempotent_replay=true`; `create_draft`/`send_draft` continuam a **1** (sem duplo agendamento/envio).
4. **T4 (AC3) — conversão/normalização UTC:** `send_at="2026-06-10T09:00:00+01:00"` → o `value` da propriedade é `"2026-06-10T08:00:00Z"` (offset aplicado). Verificar via `gc.calls` do `create_draft` **no confirm** (e `prepared["send_at_utc"]`).
5. **T5 (AC4) — validação temporal, sem token:** parametrizado:
   - `send_at` no passado → `error`, sem `confirmation_token`, `gc.calls == []` (a menos do tz best-effort).
   - `send_at` a < 2 min (ex.: now+60s) → `error`, sem token.
   - `send_at` a > 1 ano → `error`, sem token.
   - `send_at` não-parseável ("amanhã") → `error`, sem token.
   - `send_at` sem offset ("2026-06-10T09:00:00") → `error` (P3: exige offset), sem token.
6. **T6 (AC2) — sem `to`:** `schedule_prepare(to=[])` → `error`.
7. **T7 (AC9) — anexo grande diferido:** `attachments` com um `size>3MB` → `prepare` marca `large_attachments=true`; `confirm` faz `create_draft` (1) + `create_attachment_upload_session` (1) + `upload_attachment_bytes` (1) + `send_draft` (1), **nesta ordem** (verificar a ordem em `gc.calls`), e a `message` do draft leva a extended property; `send_mail`==0.
8. **T8 (AC3, degradação) — fuso indisponível:** `mailbox_timezone=None` no fake → `prepare` ok; o resumo contém a anotação de UTC; o `value` continua correto (instante absoluto do `send_at`).
9. **T9 (AC8) — auditoria:** no `confirm`, evento `action="email.schedule"`, `outcome="success"`, `recipients_count` presente, `extra` contém `large_attachments`/`send_at_utc`/`deferred:true`, `target` = `message_id`; **sem** email/assunto/corpo; **um único** `subject_hash` (o de topo).
10. **T10 (AC7) — reauth graciosa:** `auth_fail={"create_draft": 5}` (refresh falha) no `confirm` → `reauth_required`; token **não** consumido (segundo `confirm` após "re-login" volta a tentar). No `prepare`, `resolve_access_token` a falhar → `reauth_required`.
11. **T11 (P5) — aprendizagem:** com opt-in ligado, `confirm` regista `record_action_event(action="schedule")` (verificar via store/auditoria `learning.event_recorded` com `behavior_action="schedule"`); com opt-in desligado, nada.

**US-1.10 — Listar (AC1-AC6):**

12. **T12 (AC2/AC3) — listagem feliz:** `deferred_drafts` com 2 itens (1 futuro, 1 passado) → `status="ok"`; **só** o futuro vem (filtro client-side); item traz `id`, `subject` (sanitizado), `recipients_count`, `recipient_domains`, `send_at` (no fuso) e `send_at_utc`; `content_is_untrusted=true`. `gc.count("list_deferred_drafts")==1`.
13. **T13 (AC2) — query correta:** verificar em `gc.calls` que `list_deferred_drafts` foi chamado com `prop_id="SystemTime 0x3FEF"`. (No unit do client: o `$filter` testa presença e o `$expand` traz a prop — `test_graph_email_client.py`.)
14. **T14 (AC5) — sanitização:** um draft com `subject` contendo HTML/`<script>` → o `subject` devolvido vem sanitizado.
15. **T15 (AC1) — leitura não escreve:** nenhuma chamada de escrita (`create_draft`/`move_message`/`send_draft` a 0).
16. **T16 (AC6) — reauth:** `auth_fail={"list_deferred_drafts": 5}` → `reauth_required`.
17. **T17 — vazio:** sem drafts → `status="ok"`, `count=0`, `scheduled=[]`.

**US-1.11 — Cancelar (AC1-AC7):**

18. **T18 (AC1) — prepare não escreve:** `cancel_prepare(message_id)` (com `get_message` a devolver a prop) → `pending_confirmation` com token; `gc.count("move_message")==0`. (`get_message` ≤1 é leitura.)
19. **T19 (AC4) — confirm soft delete:** `confirm(token)` → `done`; `gc.count("move_message")==1` com `destination_id="deleteditems"`; `gc.count("permanent_delete")==0` (P4).
20. **T20 (AC1/idempotência):** replay → `idempotent_replay=true`; `move_message` continua a **1**.
21. **T21 (AC3) — já não é diferido:** `get_message` devolve mensagem **sem** a extended property → `cancel_prepare` devolve `error` **sem** token; `move_message`==0.
22. **T22 (AC3, degradação) — get_message falha não-auth:** se a leitura best-effort falhar (não-auth), o `prepare` ainda emite token (resumo genérico). (Modelar com o fake a devolver `{}`.)
23. **T23 (AC6) — auditoria:** `action="email.schedule_cancel"`, `target=message_id`, `extra={"permanent": false}`, sem PII, um só `subject_hash`.
24. **T24 (AC5) — reauth:** `auth_fail={"move_message":5}` no `confirm` → `reauth_required`, token não consumido.

**Transversais (herdadas, §10 da análise):** TTL (token expirado → `expired`); isolamento por subject (token de outro subject → `error`); prepare nunca emite token quando a validação falha.

### 7.3 Gates de validação MANUAL (tenant real) — **não** resolver com mocks (⬜)

Marcar em [estado-user-stories.md](estado-user-stories.md) (mesmo padrão US-1.6):

- ⬜ **Cancelabilidade real:** eliminar o rascunho antes do `send_at` impede a entrega; mapear a janela de corrida perto da hora.
- ⬜ **Localização do item** após `send_draft` (Drafts vs Outbox/transport) e quando deixa de ser cancelável.
- ⬜ **`$filter`/`$expand`** por extended property na pasta drafts fiável no tenant (senão: presença + futuro client-side, já implementado).
- ⬜ **Anexo grande (>3MB) diferido** ponta-a-ponta (herda o gate do US-1.6).
- ⬜ **Fuso:** hora apresentada vs propriedade UTC coincidem no fuso real; degradação para UTC declarada.

---

## 8. Definition of Done (checklist)

- [ ] `email.py`: `run_email_schedule_prepare`/`_confirm`, `run_email_list_scheduled`, `run_email_schedule_cancel_prepare`/`_confirm` com as assinaturas exatas de §1.
- [ ] Constante `_DEFERRED_SEND_PROP_ID = "SystemTime 0x3FEF"`, `_MIN_SCHEDULE_MARGIN = 2 min`, `_MAX_SCHEDULE_HORIZON = 365 dias`; helpers `_has_offset`/`_to_utc_iso`/`_present_in_tz`/`_extract_deferred_value`.
- [ ] **Validação temporal no `prepare`** (passado/<2min/>1ano/não-parseável/sem-offset → `error` **sem token**, **antes** de qualquer escrita).
- [ ] **`send_at` exigido com offset/Z**; normalizado para UTC na propriedade; fuso lido **só** para apresentação (best-effort, degrada para UTC declarado).
- [ ] `confirm` de US-1.9 usa **sempre** draft→send (nunca `send_mail`); a `message` do draft inclui `singleValueExtendedProperties`; anexos grandes reutilizam o caminho da US-1.6 (ordem: create_draft → upload session(s) → send_draft).
- [ ] `confirm` de US-1.11 faz **só** soft delete (`move_message`→`deleteditems`); **sem** `permanent`/`confirm_permanent`.
- [ ] `client.py`: `list_deferred_drafts` (novo) + `_map_deferred_draft`; `get_message` ganha `expand` (retrocompatível) e expõe `singleValueExtendedProperties`; `create_draft` inalterado (já serve); `move_message` reutilizado.
- [ ] `tools/_timezone.py`: `_resolve_tz` extraído; `calendar.py` e `email.py` importam de lá (sem regressão no calendário).
- [ ] `server.py`: 5 wrappers + descrições de §6; `instructions` reforçadas (recomendado).
- [ ] Auditoria: `email.schedule` (extra `{large_attachments, send_at_utc, deferred}`) e `email.schedule_cancel` (extra `{permanent:false}`), `target` correto, **sem PII**, **sem 2º subject_hash** (A1).
- [ ] Aprendizagem: `record_action_event(action="schedule")` no `confirm` de US-1.9; nada no cancelar/listar (P5).
- [ ] Listagem sanitiza assunto + `content_is_untrusted=true`; filtro "futuro" client-side.
- [ ] `FakeGraphClient` estendido (`list_deferred_drafts` + `_deferred_drafts`; `get_message` com expand/`singleValueExtendedProperties`).
- [ ] Invariantes provados por contagem; idempotência; reauth graciosa; TTL/isolamento — todos os casos de §7.2 verdes.
- [ ] `python -m pytest -q` e `python -m ruff check src tests` limpos.
- [ ] [estado-user-stories.md](estado-user-stories.md) atualizado (Implementado ☑, Testado-auto ☑); gates de §7.3 rastreados como **Validação manual ⬜**.

---

## 9. Ordem de execução Dev → QA

1. **Dev:** extrair `_resolve_tz` para `tools/_timezone.py`; ajustar `calendar.py` e correr `pytest -q` para não regredir a Fase 2.
2. **Dev:** `client.py` — `list_deferred_drafts` + `_map_deferred_draft`; `get_message` com `expand`. Unit em `test_graph_email_client.py`.
3. **Dev:** `email.py` — constantes + helpers + as 5 `run_*`. Reutilizar `_build_message`, `_att_is_large`, `_attachment_too_large`, `_domains`, `_safe_meta`, `_parse_iso`, `_confirm`.
4. **Dev:** `server.py` — 5 wrappers + descrições + instructions.
5. **Dev:** smoke local (`pytest -q`, `ruff check`).
6. **QA:** estender `FakeGraphClient` (§7.1) e escrever os casos T1-T24 (§7.2), com ênfase em: T1/T1b e T18 (prepare não escreve), T2/T7 (draft→send + ordem dos anexos grandes), T4 (UTC), T5 (validação sem token), T3/T20 (idempotência), T12/T13 (listagem + query), T19 (soft delete), T9/T23 (auditoria), T10/T16/T24 (reauth).
7. **QA:** confirmar counts a 0 nos prepares e `send_mail==0` em todo o agendamento; um só `subject_hash` nos eventos.
8. **QA:** registar os gates de §7.3 no tracking e no runbook de validação manual (cancelabilidade, localização do item, `$filter`/`$expand`, anexo grande diferido, fuso).
