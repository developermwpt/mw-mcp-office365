# Fase Email (agendamento) — estado das user stories

> **Extensão do Módulo Email** (Fase 1, US-1.1 a US-1.8) com agendamento de envio. Análise
> funcional em [analise-funcional-agendamento.md](analise-funcional-agendamento.md). Reutiliza
> os scopes **já concedidos** (`Mail.Read`/`Mail.Send`/`Mail.ReadWrite` da Fase 1 +
> `MailboxSettings.Read` da Fase 2) — **sem novo admin consent**. **Implementação Dev concluída**
> (US-1.9/1.10/1.11); testes automáticos (T1–T24) a cargo do QA.

## Legenda

- ✅ feito · ⬜ pendente
- **Implementado:** código de produção (tools + métodos Graph + registo no `server.py`).
- **Testado (automático):** coberto por testes unit/integração com Graph/Entra mockados.
- **Validação manual (tenant real):** execução no tenant/VPS reais — responsabilidade do
  cliente (ver runbook). Inclui os **gates** de §11 da análise (cancelabilidade, localização do
  item, filtro por extended property, anexo grande diferido).

## Tabela de estado

| US | Descrição | Implementado | Testado (auto) | Validação manual | Notas |
|----|-----------|:---:|:---:|:---:|-------|
| US-1.9 | Agendar envio de email para data/hora (`email_schedule_prepare`/`_confirm`) | ☑ | ☑ | ⬜ | Two-phase. Caminho **draft→send** com `singleValueExtendedProperties` `"SystemTime 0x3FEF"` (UTC ISO 8601 = `PidTagDeferredSendTime`). Hora no **fuso do mailbox** (decisão 1), convertida para UTC ao gravar. Validação: futuro + margem mínima + limite superior (recusa sem token no passado). Coexiste com anexos grandes (reutiliza o caminho da US-1.6). Auditoria `email.schedule` (só-metadados: `recipients_count`, `send_at_utc`, `large_attachments`). |
| US-1.10 | Listar envios agendados pendentes (`email_list_scheduled`) | ☑ | ☑ | ⬜ | Leitura, sem aprovação. **Fonte de verdade = o mailbox** (sem registo local): `GET /me/mailFolders/drafts/messages` com `$filter`/`$expand` sobre a extended property; "ainda futuro" filtrado client-side se necessário. Devolve `id`, assunto, contagem/domínios, `send_at` no fuso do mailbox + `send_at_utc`. Assunto/preview sanitizados + `content_is_untrusted`. |
| US-1.11 | Cancelar um envio agendado pendente (`email_schedule_cancel_prepare`/`_confirm`) | ☑ | ☑ | ⬜ | Two-phase. Cancela **eliminando o rascunho** por id — **só soft delete** (P4: move para `deleteditems`, recuperável; **sem** reforço/`permanent`). Auditoria `email.schedule_cancel` (`target=message_id`, `extra={permanent:false}`). **Cancelabilidade real é um gate** (ver Notas de validação abaixo). |

## Notas de validação manual (gates no tenant real)

Mesmo padrão do **US-1.6** da [Fase 1](../fase-1/estado-user-stories.md): os testes automáticos
(mockados) estão ✅ (T1–T24 + transversais), enquanto a validação manual no tenant permanece ⬜
e rastreada. Os 5 gates de §7.3 do contrato:

1. ⬜ **Cancelabilidade (US-1.11) — gate principal.** Confirmar que eliminar o rascunho **antes**
   do `send_at` impede a entrega; mapear a **janela de corrida** perto da hora (quando o item
   sai de Drafts para o transport e deixa de ser cancelável). A documentação oficial não cobre
   explicitamente o recall de um diferido já submetido.
2. ⬜ **Localização do item após `send_draft`.** Confirmar onde reside (Drafts até à hora vs
   Outbox/fila perto da hora) e quando deixa de ser cancelável.
3. ⬜ **Filtro/`$expand` por extended property (US-1.10).** Validar a query `$filter`/`$expand`
   sobre `"SystemTime 0x3FEF"` na pasta drafts no tenant real; se o `$filter` por valor não for
   fiável, filtrar por **presença** + "futuro" client-side (já implementado).
4. ⬜ **Anexo grande (>3MB) diferido (US-1.9).** Herda o gate do US-1.6 (envio real >3MB ainda
   ⬜): validar um agendamento com anexo grande ponta-a-ponta.
5. ⬜ **Fuso do mailbox.** Confirmar que a hora apresentada e a propriedade UTC coincidem com o
   esperado no fuso real (ex.: Hora de Lisboa) e que a degradação para UTC (sem
   `MailboxSettings.Read`) é declarada no resumo.

## Garantias transversais (a verificar por testes — herdadas da Fase 1)

- **prepare NÃO escreve:** `create_draft`/`send_draft`/`move_message` a **0** após o prepare
  (agendar e cancelar), provado por contagem.
- **Idempotência:** replay de um token consumido → `idempotent_replay=true` **sem** segundo
  draft/send/delete (não há duplo agendamento nem duplo envio).
- **TTL / isolamento:** token expirado → `expired`; token de outro `subject` → `error`.
- **Reautenticação graciosa:** falha de refresh → `reauth_required`; no `confirm` o token
  **não é consumido** (repetível após re-login). A leitura do fuso é best-effort e nunca
  derruba a sessão.
- **Auditoria só-metadados:** `email.schedule` / `email.schedule_cancel` com `subject_hash` de
  topo, contagem de destinatários e extras seguros (`send_at_utc`, `large_attachments`,
  `permanent`) — nunca o corpo, o assunto em claro nem endereços.
- **Sanitização:** assunto/preview da listagem sanitizados; `content_is_untrusted=true`.

## Onde estão os testes

- **Unit do GraphClient** — `tests/unit/test_graph_email_client.py`:
  `test_list_deferred_drafts_query_filter_expand_select` e `_sem_prop_mapeia_none` (query
  `$filter`/`$expand`/`$select`/`$top` e mapeamento), `test_get_message_com_expand_expoe_extended_property`
  e `_sem_expand_retrocompativel` (P7, retrocompatível).
- **Integração — agendar/cancelar + transversais** — `tests/integration/test_email_write_e2e.py`:
  T1/T1b (prepare não escreve), T2 (draft→send com a propriedade), T3 (idempotência), T4 (UTC),
  T5 (validação temporal sem token, parametrizado: passado/<2min/>1ano/não-parseável/sem-offset),
  T6 (sem `to`), T7 (anexo grande: ordem + propriedade), T8 (fuso indisponível → UTC), T9
  (auditoria `email.schedule`), T10/T10b (reauth no prepare e no confirm, token não consumido),
  **T11/T11b** (aprendizagem `action="schedule"` registada com opt-in; nada sem opt-in); T18 (cancel prepare não
  escreve), T19 (soft delete), T20 (idempotência), T21 (já não diferido → error sem token), T22
  (degradação de `get_message`), T23 (auditoria `email.schedule_cancel`), T24 (reauth, token não
  consumido); transversais TTL (`expired`) e isolamento por subject (`error`).
- **Integração — listar** — `tests/integration/test_email_read_e2e.py`: T12 (filtro "futuro"
  client-side + item completo), T13 (query com `prop_id`), T14 (sanitização do assunto), T15
  (leitura não escreve), T16 (reauth), T17 (vazio).
- **Infra de teste** — `tests/integration/fake_graph.py`: `list_deferred_drafts` + atributo
  programável `_deferred_drafts`; `get_message` tolera `expand` e o `_message` programável pode
  trazer `singleValueExtendedProperties`.

> **Bug corrigido (T11 / P5):** o QA apanhou que `"schedule"` não constava de
> `_LEARNABLE_ACTIONS` (`src/mcp_o365/learning/events.py`), pelo que `build_behavior_event`
> devolvia `None` e o `confirm` de US-1.9 **não** registava o evento de aprendizagem exigido
> pelo contrato P5. Corrigido (acrescentado `"schedule"` ao conjunto); `test_schedule_aprendizagem_action_schedule_T11`
> passa agora normalmente. O recommender ignora ações fora do seu mapa (`"schedule"`, tal como
> `"send"`, é aprendível mas não sugerido) — sem regressão.

Correr: `python -m pytest -q` · lint: `python -m ruff check src tests`.
