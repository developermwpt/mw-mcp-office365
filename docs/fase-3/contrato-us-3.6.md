# Contrato de Implementação — US-3.6 (barreira server-side anti-fuga no envio a contacto explícito)

**Projeto:** mw-mcp-office365 · **Fase:** 3 (Teams) · **User story:** US-3.6 (D12)
**Documento:** Contrato técnico para o Dev/QA. Decisão fechada pelo coordenador: **D12 = B + C + A**.
**Estado:** Pronto para implementação. Análise funcional fechada (ver [analise-funcional-teams.md](analise-funcional-teams.md) §US-3.6/D12 e [estado-user-stories.md](estado-user-stories.md) §US-3.6 — 9 critérios de aceitação).
**Precedências:** US-3.3/3.4/3.5 já implementadas; `run_teams_send_message_prepare` já exclui o próprio emissor do cálculo de `others` (commit 61b7cde). Este contrato adiciona **apenas** a barreira `intended_recipient`. **Não** altera o `confirm`, o `get_chat`, nem o `get_or_create_one_on_one_chat`.

> **Âmbito.** B (núcleo): guarda server-side aplicável no `prepare`. C: o caminho fácil seguro (`get_or_create_one_on_one_chat`) já faz só match `oneOnOne` — **mantém-se inalterado**. A: endurecer descrições das tools no `server.py`. O Dev segue este contrato à risca; **não** se escreve código aqui (este é doc). Testes com Graph mockado (FakeGraphClient já tem `me`/`get_chat`/`send_chat_message`).

---

## 1. Núcleo (B) — alteração de `run_teams_send_message_prepare`

### 1.1 Assinatura (exata, em `src/mcp_o365/tools/teams.py`)

Adiciona-se **um** parâmetro keyword-only opcional, **a seguir a `body_type`** (mantendo a ordem dos restantes; é opcional e default `None` → retrocompatível):

```python
async def run_teams_send_message_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    chat_id: str,
    body: str,
    body_type: str = "text",
    intended_recipient: str | None = None,   # <-- NOVO (US-3.6/D12)
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
```

`intended_recipient`: email do contacto explicitamente nomeado pelo utilizador. Se `None` (ausente), o comportamento atual mantém-se **integralmente** (incl. a degradação que emite token quando o `get_chat` falha). Se presente, ativa a barreira da §1.3.

A docstring ganha uma frase no fim:
> *US-3.6/D12 — Se `intended_recipient` for indicado, EXIGE que o `chat_id` seja um `oneOnOne` cujo único outro membro (excluído o próprio) == `intended_recipient` (case-insensitive); caso contrário recusa com `error` SEM emitir token (e regista auditoria `teams.send_blocked`). Fail-closed: se o `get_chat` degradar e não houver `intended_recipient` mantém-se o resumo degradado com token; se houver `intended_recipient`, recusa (não dá para verificar o destino).*

### 1.2 Onde entra no fluxo

A barreira corre **depois** do bloco que já existe (resolução do próprio via `/me` best-effort e cálculo de `others`/`other_emails` — linhas ~283-296 do `teams.py` atual) e **antes** da montagem do `summary`/`approval.prepare(...)`. Reaproveita `chat_type`, `own_cf`, `others` e `other_emails` já calculados. **Não** se faz nenhuma chamada Graph nova (a verificação usa o `chat` e o `me` já lidos best-effort no início do prepare).

### 1.3 Pseudocódigo (fiel ao estilo do `teams.py`; inserir entre o cálculo de `other_emails` e o `if chat_type == "oneOnOne":`)

```python
    # US-3.6/D12 — Barreira anti-fuga (B). Quando o utilizador nomeia uma pessoa, o servidor
    # EXIGE que o destino seja o 1:1 exato com ela; nunca um grupo nem outra pessoa. Fail-closed:
    # se o get_chat degradou (chat_type indisponível) não dá para verificar -> recusa.
    if intended_recipient is not None:
        target_cf = intended_recipient.casefold()
        others_cf = [(m.get("email") or "").casefold() for m in others]
        is_strict_one_on_one = (
            chat_type == "oneOnOne"
            and others_cf == [target_cf]   # ÚNICO outro membro == alvo (já exclui o próprio)
        )
        if not is_strict_one_on_one:
            log_audit(
                audit_logger, action="teams.send_blocked", subject=subject,
                account_id=account.account_id, target=chat_id, outcome="blocked",
                extra={"reason": "intended_recipient_mismatch", "chat_type": chat_type},
            )
            return {
                "status": "error",
                "message": (
                    "Por segurança, não enviei a mensagem: pediu para enviar a uma pessoa "
                    "específica, mas este chat não é a conversa 1:1 com essa pessoa (pode ser "
                    "um grupo, outra conversa, ou não foi possível confirmar o destinatário). "
                    "Para enviar a essa pessoa, obtenha primeiro a conversa 1:1 com "
                    "`teams_get_or_create_one_on_one_chat_prepare` e use o `chat_id` devolvido."
                ),
            }
```

**Notas de fidelidade:**
- A comparação `others_cf == [target_cf]` é a **mesma** lógica de `_find_one_on_one` (lista de outros emails em casefold, comparada com `[target]`). Cobre de uma vez: grupo (`chat_type != "oneOnOne"` → falha), 1:1 com outra pessoa (`others_cf != [target_cf]`), membro alvo sem email (entra como `""`, nunca == `target_cf` → falha), e get_chat degradado (`chat_type is None` → falha = **fail-closed**, diretiva 1 do coordenador).
- O `account` já existe nesta altura (resolvido em `resolve_access_token`, ~linha 258), por isso `account.account_id` está disponível para a auditoria.
- A barreira devolve `{"status": "error", ...}` **sem** chamar `approval.prepare(...)` → **sem `confirmation_token`**. Sem token, o `confirm` nada pode enviar (critério 9).
- O `intended_recipient` **não** entra no `payload` da aprovação (não é necessário no confirm e evita guardar PII no token).

### 1.4 Comportamento quando `intended_recipient is None`

Inalterado: o fluxo segue para o `if chat_type == "oneOnOne": ... elif chat_type: ... else: (degradado, emite token)` exatamente como hoje. Envio a grupo continua possível com confirmação humana (critério 5); 1:1 com a pessoa continua a emitir token (critério 4).

---

## 2. Auditoria da recusa (diretiva 2)

Assinatura **real** confirmada em `src/mcp_o365/observability/audit.py`:

```python
def log_audit(logger, *, action, subject, account_id=None, target=None,
              outcome, recipients_count=None, extra=None) -> None
```

`extra` é fundido por `fields.update(extra)` **depois** de o `subject_hash` de topo (identidade) já estar escrito. Por isso (regra A1 do projeto): **NÃO** colocar `subject_hash` nem qualquer email/nome em `extra` (sobrescreveria o hash de identidade). Chamada exata da recusa:

```python
log_audit(
    audit_logger, action="teams.send_blocked", subject=subject,
    account_id=account.account_id, target=chat_id, outcome="blocked",
    extra={"reason": "intended_recipient_mismatch", "chat_type": chat_type},
)
```

- `action="teams.send_blocked"` (novo); `outcome="blocked"` (novo outcome, distinto de success/error).
- `target=chat_id` (não é PII; é o mesmo padrão que `teams.send`).
- `extra`: só metadados — `reason` (string-constante) + `chat_type` (`"group"`/`"oneOnOne"`/`None`). **Sem** emails, nomes, texto da mensagem, nem `intended_recipient`. **Sem** segundo `subject_hash`.
- `audit_logger` já existe no módulo (`logging.getLogger("mcp_o365.audit")`, linha 51).

---

## 3. Descrições e schema das tools no `server.py` (A)

O `@mcp.tool(description=...)` é a fonte da descrição apresentada ao LLM; a **assinatura da função wrapper** é o schema que o LLM vê. Mudam **duas** coisas: (a) o wrapper de `teams_send_message_prepare` ganha o parâmetro `intended_recipient` (passa a fazer parte do schema MCP exposto ao LLM); (b) endurecem-se duas descrições.

### 3.1 Mudança no schema MCP — wrapper `teams_send_message_prepare` (linhas ~638-644)

**Atual:**
```python
    async def teams_send_message_prepare(
        chat_id: str, body: str, body_type: str = "text"
    ) -> dict:
        return await teams_tools.run_teams_send_message_prepare(
            _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
            store=store, approval=approval, chat_id=chat_id, body=body, body_type=body_type,
        )
```

**Novo:**
```python
    async def teams_send_message_prepare(
        chat_id: str, body: str, body_type: str = "text",
        intended_recipient: str | None = None,
    ) -> dict:
        return await teams_tools.run_teams_send_message_prepare(
            _subject(), mapping=mapping, plane_b=plane_b, graph_client=graph_client,
            store=store, approval=approval, chat_id=chat_id, body=body, body_type=body_type,
            intended_recipient=intended_recipient,
        )
```

Efeito no schema: o LLM passa a ver um campo opcional `intended_recipient` (string, nullable). É a única mudança de schema da US-3.6.

### 3.2 Descrição de `teams_send_message_prepare` (linhas ~626-636)

**Atual:**
> "FASE 1/2 — Prepara o envio de uma mensagem para um chat de Teams EXISTENTE (NÃO envia). Também serve para RESPONDER numa conversa (em chats não há thread: responder = enviar no mesmo chat). Parâmetros: `chat_id` (de `teams_list_chats` ou de `teams_get_or_create_one_on_one_chat_*`), `body`, `body_type` ('text' por defeito; 'html' só se o utilizador pedir formatação). Valida o tamanho (máximo ~28000 caracteres) e devolve um resumo + `confirmation_token`. O resumo declara o tipo de chat, quantos participantes e em que domínios. Chame `teams_send_message_confirm`."

**Novo (acrescenta o bloco de segurança sobre `intended_recipient`; resto igual):**
> "FASE 1/2 — Prepara o envio de uma mensagem para um chat de Teams EXISTENTE (NÃO envia). Também serve para RESPONDER numa conversa (em chats não há thread: responder = enviar no mesmo chat). Parâmetros: `chat_id` (de `teams_list_chats` ou de `teams_get_or_create_one_on_one_chat_*`), `body`, `body_type` ('text' por defeito; 'html' só se o utilizador pedir formatação), e `intended_recipient` (opcional). **SEGURANÇA — quando o utilizador pede para enviar a uma PESSOA NOMEADA (por nome/email): obtenha o `chat_id` via `teams_get_or_create_one_on_one_chat_prepare` e passe SEMPRE `intended_recipient` com o email dessa pessoa.** Se `intended_recipient` for indicado, o servidor RECUSA o envio (sem token) caso o `chat_id` não seja a conversa 1:1 exata com essa pessoa — nunca use um `chat_id` de grupo vindo de uma pesquisa por nome para enviar a uma pessoa. Para enviar mesmo a um GRUPO, NÃO passe `intended_recipient`. Valida o tamanho (máximo ~28000 caracteres) e devolve um resumo + `confirmation_token`. O resumo declara o tipo de chat, quantos participantes e em que domínios. Chame `teams_send_message_confirm`."

### 3.3 Descrição de `teams_list_chats` (linhas ~583-594) — reforço leve

**Atual (frase final):**
> "… Para enviar a uma PESSOA por nome, NÃO adivinhe o chat: use `resolve_recipient` e depois `teams_get_or_create_one_on_one_chat_prepare`."

**Novo (frase final):**
> "… Para enviar a uma PESSOA por nome, NÃO adivinhe o chat nem use um `chat_id` de grupo desta lista: use `resolve_recipient` e depois `teams_get_or_create_one_on_one_chat_prepare`, e ao enviar passe `intended_recipient` com o email da pessoa (o servidor recusa o envio se o chat não for o 1:1 exato)."

### 3.4 `instructions` do servidor (linhas ~84-92) — opcional, recomendado

A nota global já encaminha "manda à X" por resolver→1:1. Acrescentar, no fim dessa frase, "… e ao enviar a uma pessoa nomeada passar sempre `intended_recipient`." Não é bloqueante (a garantia é a barreira B); é reforço A.

> O `confirm` (`teams_send_message_confirm`) e a sua descrição **não mudam**.

---

## 4. Estratégia de teste (QA)

**Ficheiro:** estender `tests/integration/test_teams_write_e2e.py`. **Mocks:** já existentes no `FakeGraphClient` — `me()` devolve `self._me` (resolução do próprio email via `userPrincipalName`); `get_chat()` devolve `self._chat`; `send_chat_message()` registado por `_record`. Contagens via `count("send_chat_message")` e `count("create_one_on_one_chat")`. Auditoria via captura do logger `mcp_o365.audit` (mesmo padrão dos testes de `teams.send`/`email.send`).

**Setup base do `_chat`/`_me`:**
- `_me = {"userPrincipalName": "eu@empresa.pt", ...}` (o próprio).
- 1:1 com vera: `_chat = {"chat_type": "oneOnOne", "members": [{"email": "eu@empresa.pt", ...}, {"email": "vera@empresa.pt", "name": "Vera"}], ...}`.
- grupo com vera: `_chat = {"chat_type": "group", "topic": "Projeto X", "members": [eu, vera, joao], ...}`.
- 1:1 com outra: `_chat = {"chat_type": "oneOnOne", "members": [eu, {"email": "ana@empresa.pt"}]}`.
- 1:1 com vera sem email no alvo (A5): membro alvo só com `aad_user_id`, sem `email`.
- get_chat degradado: `auth_fail`/exceção não-auth no `get_chat` → `chat = {}` → `chat_type is None`.

**Casos (mapeados aos 9 critérios do Analista):**

1. **C1 — 1:1 correto + recipient:** `prepare(chat_id=1:1_vera, body, intended_recipient="vera@empresa.pt")` → `status` de `pending_confirmation` **com** `confirmation_token`; `count("send_chat_message")==0`.
2. **C2 — grupo + recipient (vetor principal):** `prepare(chat_id=grupo_com_vera, body, intended_recipient="vera@empresa.pt")` → `status="error"`, **sem** `confirmation_token`; `count("send_chat_message")==0`; `message` menciona a conversa 1:1 / `teams_get_or_create_one_on_one_chat_prepare`.
3. **C3 — 1:1 com outra pessoa + recipient:** `prepare(chat_id=1:1_ana, body, intended_recipient="vera@empresa.pt")` → `error`, sem token (chat_id trocado de pessoa).
4. **C4 — 1:1 correto, recipient ausente:** `prepare(chat_id=1:1_vera, body)` (sem `intended_recipient`) → `pending_confirmation` com token (retrocompatível).
5. **C5 — grupo, recipient ausente:** `prepare(chat_id=grupo, body)` → `pending_confirmation` com token (envio a grupo legítimo continua a funcionar).
6. **C6 — pessoa só em grupos:** com `intended_recipient` e `chat_id` de grupo (não há 1:1) → recusa (é o C2 reforçado; a mensagem orienta a criar o 1:1).
7. **C7 — alvo sem email (A5):** 1:1 cujo outro membro só tem `aad_user_id` + `intended_recipient="vera@…"` → `error`, sem token (degradação segura: `""` != alvo).
8. **C8 — case-insensitive + exclusão do próprio:** `prepare(chat_id=1:1_vera, body, intended_recipient="VERA@EMPRESA.PT")` → `pending_confirmation` com token (match case-insensitive); confirmar que o próprio (`eu@empresa.pt`) presente em `members` **não** quebra o `others_cf == [target]`.
9. **C9 — recusa não toca tokens nem Graph de escrita:** em qualquer recusa (C2/C3/C6/C7), `count("send_chat_message")==0`, **nenhum** `confirmation_token` na resposta, e (se o teste verificar o store) nenhum token criado/consumido.
10. **C10 — FAIL-CLOSED (diretiva 1):** `get_chat` degradado (não-auth) **com** `intended_recipient="vera@…"` → `error`, sem token (chat_type=None → não verificável → recusa). Contraste: mesmo `get_chat` degradado **sem** `intended_recipient` → `pending_confirmation` com token (resumo degradado, comportamento atual — já coberto na US-3.3, reconfirmar aqui).
11. **C11 — AUDITORIA blocked (diretiva 2):** numa recusa (C2), o logger `mcp_o365.audit` emite **um** evento com `action="teams.send_blocked"`, `outcome="blocked"`, `target==chat_id`, `extra` contém `reason` + `chat_type` e **NÃO** contém emails/nomes/texto nem um segundo `subject_hash`; o `subject_hash` de topo (identidade) está presente uma única vez.

**Garantias transversais a reconfirmar:** o `confirm` não muda; idempotência do token intacta; nenhuma chamada Graph nova no caminho da barreira (a verificação reaproveita `chat`/`me` já lidos — `count("get_chat")` continua ≤1 no prepare, `count("me")` ≤1).

---

## 5. Definition of Done (US-3.6)

- [ ] `run_teams_send_message_prepare` aceita `intended_recipient: str | None = None`; barreira inserida **após** o cálculo de `others`/`other_emails` e **antes** do `summary`/`approval.prepare`.
- [ ] Com `intended_recipient` presente: emite token **apenas** se `chat_type == "oneOnOne"` e `others_cf == [intended_recipient.casefold()]`; caso contrário `{"status":"error", ...}` **sem** `confirmation_token` e **sem** chamar `approval.prepare`.
- [ ] Fail-closed coberto: `get_chat` degradado + `intended_recipient` → recusa (chat_type=None); sem `intended_recipient` → mantém token degradado.
- [ ] Recusa audita `teams.send_blocked`/`outcome="blocked"`, `target=chat_id`, `extra={reason, chat_type}` — **sem** PII e **sem** segundo `subject_hash` (A1).
- [ ] Mensagem de erro PT-PT da §1.3 (clara, orientadora, sem expor dados sensíveis).
- [ ] `server.py`: wrapper de `teams_send_message_prepare` ganha `intended_recipient` (schema MCP); descrições de `teams_send_message_prepare` e `teams_list_chats` endurecidas (§3.2/§3.3); `instructions` reforçadas (§3.4, opcional).
- [ ] `confirm`, `get_chat`, `get_or_create_one_on_one_chat`: **inalterados** (C mantém-se).
- [ ] Retrocompatibilidade: chamadas sem `intended_recipient` comportam-se exatamente como hoje (C4/C5 verdes).
- [ ] 11 casos de teste (§4) verdes com Graph mockado; `python -m pytest -q` e `python -m ruff check src tests` limpos.
- [ ] Atualizar [estado-user-stories.md](estado-user-stories.md) §US-3.6 (Dev ☑, QA ☑) e marcar D12 como fechado.

## 6. Ordem de execução Dev → QA

1. **Dev:** alterar `teams.py` (assinatura + barreira + auditoria — §1, §2). Não tocar no `confirm`.
2. **Dev:** alterar `server.py` (schema do wrapper + 2 descrições + instructions — §3).
3. **Dev:** smoke local (`pytest -q` para não regredir US-3.3/3.4/3.5).
4. **QA:** escrever/estender os 11 casos (§4), com ênfase em C2 (vetor principal), C10 (fail-closed) e C11 (auditoria blocked). Confirmar counts (`send_chat_message==0` em toda recusa).
5. **QA:** validar retrocompatibilidade (C4/C5) e a ausência de chamadas Graph extra no caminho da barreira.
6. **QA:** atualizar o tracking e o [runbook-validacao-manual.md](runbook-validacao-manual.md) com um passo manual: "enviar à Vera com `intended_recipient` apontando um `chat_id` de grupo → tem de recusar sem enviar".
