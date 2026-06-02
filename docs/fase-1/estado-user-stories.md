# Fase 1 — Módulo Email: estado das user stories

> **Gate G3 ✅ validado (2026-06-02).** O refresh do token Graph **sob Conditional Access**
> (dispositivo gerido) foi testado no tenant real e **manteve acesso** — o refresh silencioso
> do servidor não é bloqueado pela CA atual. O bloqueador central da PoC está ultrapassado;
> não é preciso exceção de CA (named location). Admin consent concedido para
> `Mail.Read`/`Mail.Send`/`Mail.ReadWrite`.

## Legenda

- ✅ feito · ⬜ pendente
- **Testado (automático):** coberto por testes unit/integração com Graph/Entra mockados.
- **Validação manual (tenant real):** execução no tenant/VPS reais — responsabilidade do
  cliente (ver runbook). Pendente em todas as US enquanto G3 não estiver validado.

## Tabela de estado

| US | Descrição curta | Implementado | Testado (auto) | Validação manual | Notas |
|----|-----------------|:---:|:---:|:---:|-------|
| US-1.1 | Pesquisar emails (`$search`/`$filter`, paginação `has_more`) | ✅ | ✅ | ⬜ | `$search` envia `ConsistencyLevel: eventual`. |
| US-1.2 | Ler email (corpo sanitizado, `content_is_untrusted`) | ✅ | ✅ | ⬜ | HTML sanitizado (anti prompt injection); flag de não-confiança sempre presente. |
| US-1.3 | Enviar email (prepare/confirm) | ✅ | ✅ | ⬜ | Two-phase approval + auditoria `email.send`. |
| US-1.4 | Responder / responder-a-todos / reencaminhar | ✅ | ✅ | ⬜ | Mantém a thread; forward exige `to_recipients`; auditoria `email.reply`/`email.forward`. |
| US-1.5 | Listar e descarregar anexos (`contentBytes`) | ✅ | ✅ | ⬜ | Leitura, sem aprovação. |
| US-1.6 | Anexos grandes (>3MB) no envio | ✅ | ✅ | ⬜ | Upload session completo (rascunho + chunks + envio). |
| US-1.7 | Mover email entre pastas | ✅ | ✅ | ⬜ | Resolve nome de pasta → id (bem-conhecidas + `list_folders`); auditoria `email.move`. |
| US-1.8 | Eliminar email (soft + permanente reforçada) | ✅ | ✅ | ⬜ | Permanente recusada sem `confirm_permanent=True` (antes de consumir o token); auditoria `email.delete`. |

## US-1.6 — anexos > 3 MB (completo)

Anexos acima de 3 MB não podem ir inline no `POST /me/sendMail`. O fluxo está implementado
ponta-a-ponta e testado: o `send_prepare` marca `large_attachments=true` e o `send_confirm`
segue o caminho de **rascunho** — `create_draft` (com os anexos inline ≤3MB) +
`create_attachment_upload_session` (por anexo grande) + **`upload_attachment_bytes`** (PUT dos
bytes em chunks de 320 KiB com `Content-Range`, na `uploadUrl` pré-autenticada, sem `Bearer`)
+ `send_draft`. Coberto por `test_upload_attachment_bytes_em_chunks_sem_bearer` e pelo E2E
`test_send_anexo_grande_*`.

## Garantias transversais (verificadas por testes)

- **Aprovação em duas fases:** `prepare` devolve token + resumo + `expires_at` e **não toca
  no Graph**; `confirm` resolve um token Graph **fresco** e só então executa.
- **Idempotência:** reapresentar um token já consumido devolve `idempotent_replay=true` **sem
  re-executar** (não duplica envios/eliminações).
- **TTL / isolamento:** token expirado → `expired`; token de outro `subject` → `error`
  (`ConfirmationNotFound`).
- **Reautenticação graciosa:** qualquer falha de refresh (ex.: `invalid_grant` da CA) →
  `{"status":"reauth_required"}` e o Graph **não é chamado**.
- **Auditoria sem PII:** cada escrita emite `event=audit` com `subject_hash` (nunca o subject
  em claro) e, no máximo, contagem de destinatários — nunca endereços nem corpo.
- **Sanitização:** `<script>`/`<style>`, handlers `on*`, `javascript:` URIs, conteúdo oculto
  (`display:none`/`hidden`) e comentários são removidos do corpo recebido.

## Onde estão os testes

- Unit: `tests/unit/test_approval_engine.py`, `test_session_helper.py`, `test_sanitize.py`,
  `test_audit.py`, `test_graph_email_client.py`.
- Integração (tools ponta-a-ponta): `tests/integration/test_email_read_e2e.py`,
  `test_email_write_e2e.py` (com `tests/integration/fake_graph.py`).

Correr: `python -m pytest -q` · lint: `python -m ruff check src tests`.
