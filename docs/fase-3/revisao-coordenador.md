# Revisão do Coordenador Técnico — Fase 3 (Teams / chats)

**Data:** 2026-06-06
**Documento revisto:** [`docs/fase-3/plano-implementacao.md`](plano-implementacao.md) (contrato técnico do PM, decisões D1–D11)
**Documento de apoio:** [`docs/fase-3/analise-funcional-teams.md`](analise-funcional-teams.md)
**Revisor:** Coordenador Técnico (revisão crítica independente — gate de aprovação pré-implementação)
**Fontes de verdade consultadas:** `src/mcp_o365/graph/client.py`, `tools/calendar.py`, `tools/email.py`, `tools/contacts.py`, `tools/_session.py`, `approval/engine.py`, `observability/audit.py`, `logging_setup.py`, `config.py`, `tests/integration/fake_graph.py`.

---

## Veredicto global: **APROVADO COM RESSALVAS**

O plano é sólido, coerente com os padrões da Fase 1/2 e maduro na parte de segurança. Reutiliza fielmente o dual-plane (`resolve_access_token`/`call_graph`), o `ApprovalEngine` (idempotência por token), a reauth graciosa, a fronteira de sanitização na tool e a auditoria só-metadados. As decisões D1–D11 são, na generalidade, as mais simples e seguras.

**Não há bloqueadores que impeçam o developer de arrancar** na ordem proposta (§7.1: scopes + client + mapeadores + US-3.1). Existem, contudo, **dois achados Maiores** que têm de ser resolvidos no plano/contrato **antes** de fechar as escritas (US-3.3/3.4) — ver A1 e A2 — e que listo como pré-condições. Os restantes são Menores/Notas que podem ser tratados durante a implementação.

---

## Pontos fortes

1. **Reutilização disciplinada dos padrões.** A separação client (corpo CRU) ↔ tool (sanitização + `content_is_untrusted`) está corretamente herdada do email/calendário; o `_request` real já suporta tudo o que os novos métodos pedem (URLs absolutos para `@odata.nextLink` — linha 535 de `client.py` —, headers, retry 429, 401/403→`UpstreamAuthError`). Nenhum método novo obriga a tocar no núcleo HTTP. Confirmado.
2. **Idempotência bem fundamentada.** O `ApprovalEngine.confirm` (engine.py L88-90) devolve `idempotent_replay` sem re-executar; aplicar isto a `send_chat_message`/`create_one_on_one_chat` neutraliza o risco real de duplicação num chat. A estratégia de prova por contagem de chamadas (FakeGraphClient.count) é a correta e está alinhada com a Fase 1/2.
3. **D3 (procurar→criar) bem desenhado.** Resolver o próprio email via `/me` (reutilizando `_own_email` de `calendar.py`), match por email case-insensitive excluindo o próprio, devolução de `ok` SEM token quando o chat existe (criação a 0) — fluxo limpo e testável.
4. **Fronteira de conteúdo não-confiável traçada onde deve.** D8 (`is_system = messageType != "message"`, cartões só como metadados) e a sanitização do `last_message_preview`/`body` na tool estão corretas e cobrem o vetor de prompt injection (R3).
5. **D9 alinhado com o Calendário.** `resolve_recipient` a montante; tools só aceitam emails/`chat_id` resolvidos; reforço em `instructions` e playbook — exatamente o padrão validado na Fase 2.

---

## Achados (por severidade)

### Bloqueadores
Nenhum.

### Maiores

**A1 (Maior) — Colisão semântica do `subject_hash` na auditoria `teams.send`. Ref.: §5 ponto 8 + MICRO-DECISÃO; §8.**
O `log_audit` **já emite** `"subject_hash": subject_hash(subject)` onde `subject` é a **identidade do utilizador** (audit.py L41; helper em logging_setup.py L27-29 — pseudonimização RGPD do *quem*). A micro-decisão do plano propõe colocar, dentro de `extra`, um **segundo** `subject_hash` calculado sobre o `chat_id`. Como `log_audit` faz `fields.update(extra)` (audit.py L49), o `extra["subject_hash"]` **sobrescreveria silenciosamente o hash da identidade do utilizador** no registo de auditoria — uma regressão de privacidade/rastreabilidade, não uma melhoria.
- Note-se ainda que o `chat_id` do Graph **não é PII** (é um identificador opaco) e o plano já o regista como `target` (§5.8) — logo "hashá-lo" não acrescenta proteção; só duplica/colide.
- O precedente real é claro: `email.send` (email.py L509-517) **não** põe nenhum `subject_hash` em `extra` (só `large_attachments`); o `subject_hash` da identidade vem sempre do `log_audit`. O calendário usa `subject_hash(event_subject)` porque o *assunto do evento* É conteúdo do utilizador — mas uma mensagem de chat **não tem assunto**, como o próprio plano admite.
- **Recomendação (bloqueante para US-3.3):** **remover** o `subject_hash` de `extra` em `teams.send`. O `target=chat_id` já identifica o recurso sem PII; o `subject_hash` da identidade já vem do `log_audit`. `extra` de `teams.send` deve conter apenas `{chat_type, body_type}`. Atualizar §5.8 e §8 em conformidade. (A análise funcional §4.6 dizia "se aplicável, `subject_hash` de uma referência curta" — fica resolvido por "não aplicável".)

**A2 (Maior) — Resumo do `teams_send_message_prepare` pode degradar com demasiada frequência por causa do `list_chats(top=50)`. Ref.: §3 US-3.3 "Leitura acessória"; §2.1 `list_chats`.**
O prepare do envio resume "tipo de chat + N participantes + domínios" lendo o chat via `list_chats` + match por `chat_id`. Mas `list_chats` traz **apenas a 1ª página** (`top=50`, sem paginar no caminho do envio). Um utilizador com mais de 50 chats que envie para um chat fora da 1ª página obtém **sempre** um resumo sem detalhes (degradação graciosa) — o que enfraquece o valor de confirmação humana exatamente na operação de escrita mais frequente. Pior: o resumo é a barreira anti-erro ("enviou para o grupo errado").
- **Recomendação (bloqueante para US-3.3):** acrescentar um método `get_chat(chat_id)` em `client.py` (`GET /me/chats/{id}?$expand=members`, read) e usá-lo no prepare em vez do match sobre `list_chats`. É mais barato, determinístico e robusto a >50 chats. Mantém-se best-effort (falha não-auth → resumo sem detalhes; `ReauthRequired` → `reauth_response`), igual ao `_resolve_tz`. Se o PM preferir não acrescentar método, então o plano tem de assumir explicitamente que o resumo do envio pode não ter detalhes do chat e o `instructions`/playbook devem instruir o LLM a confirmar o chat por outra via — solução inferior. Recomendo o `get_chat`.

### Menores

**A3 (Menor) — Inconsistência de endpoint `/me/chats/{id}/messages` (POST) entre plano e análise/Graph. Ref.: §2.1 `send_chat_message`; nota §2.1 "/me/chats vs /chats".**
A análise funcional (§3, tabela) e a documentação do Graph usam `POST /chats/{id}/messages` (sem `/me`); o plano padroniza em `/me/chats/{id}/messages` por coerência. Em delegated ambos funcionam para chats do próprio, mas o `POST .../messages` sob `/me/` é menos comummente documentado. **Recomendação:** aceitável como decisão de coerência, MAS o developer deve **validar no tenant real** (runbook) que o POST sob `/me/chats/...` aceita o envio; se devolver erro, cair para `/chats/{id}/messages`. Registar no runbook como ponto de verificação explícito. Não bloqueia os testes mockados.

**A4 (Menor) — `_chat_from` e o fallback de email do membro: lógica frágil/ambígua. Ref.: §4.2 nota `_chat_from`; §2.1 `_map_chat_member`; §4.1.**
A expressão proposta `user.get("email") or user.get("userIdentityType") and None` (§4.2) é confusa e quase de certeza não faz o que aparenta (devolve sempre `None` ou o email, nunca o `userIdentityType`). O `from.user` do Graph normalmente **não traz email** — só `id` e `displayName`. Há ainda incoerência entre `_map_chat_member` (§2.1: expõe `aad_user_id` no fallback) e `_map_chat_summary` (§4.1: já inclui `aad_user_id` sempre) — há dois contratos diferentes para o mesmo conceito. **Recomendação:** simplificar `_chat_from` para `{"name": user.get("displayName"), "email": user.get("email")}` (email tipicamente `None` em mensagens — aceitável; é só para apresentação), e `None` quando `from`/`from.user` for nulo. Unificar o mapeamento de membro num único helper. Tratar na implementação; cobrir com os unit tests já previstos em §6.3.

**A5 (Menor) — `D3`/matching 1:1 quando o membro só tem `userId` (sem email). Ref.: §3 US-3.4 "D3 — procurar primeiro"; R2.**
O match do 1:1 existente é por email; se os `members` do chat só trouxerem `userId` (cenário real — o `email` em `aadUserConversationMember` nem sempre vem), o match **falha** e o fluxo segue para criar. Como o `POST /chats` é idempotente no Graph, isto não duplica o chat — mas faz uma escrita confirmada (pede confirmação ao utilizador) que podia ter sido evitada. O plano reconhece-o em R2 ("match falha graciosamente e segue-se para criação"), o que é aceitável para a v1. **Recomendação:** manter, mas documentar na DoD da US-3.4 que "membro sem email → não há match → cria (idempotente)" é comportamento esperado, e cobrir esse caso num teste (evita falso-positivo de regressão). Considerar, como melhoria futura, resolver o `userId` próprio via `/me` (já feito) e comparar por `userId` quando o email faltar.

**A6 (Menor) — `call_graph` repete o `op` uma vez em `UpstreamAuthError`: risco teórico de duplo-POST no envio. Ref.: §5 ponto 6; `_session.py` L127-141.**
`call_graph` força refresh e repete `op` uma vez se o Graph devolver 401/403. Se um `POST .../messages` for aceite pelo Graph mas a resposta surgir como 401/403 (cenário muito improvável), o retry reenviaria. Isto é **exatamente** o comportamento já existente em `send_mail`/`create_event` (Fase 1/2), portanto é coerente e não é uma regressão. A camada de idempotência (token) protege contra duplicação ao nível do MCP/LLM, não ao nível de um retry interno do `call_graph`. **Recomendação:** Nota apenas — aceitar o risco residual como na Fase 1/2; não exigir mudança. Mencionar em §7.2 (R4) para ficar explícito que a idempotência cobre o replay do token, não o retry-auth interno.

**A7 (Menor) — Paginação da listagem sem `next_link` exposto ao cliente (MICRO-DECISÃO §3 US-3.1). Ref.: §3 US-3.1; D5.**
A listagem traz até `_MAX_LIST_FETCH=200` e não pagina para o cliente; difere conscientemente de D5 (mensagens) que expõe `next_link`. É coerente (volume de chats é baixo; filtro client-side) e simples. **Recomendação:** aceitável. Acrescentar à DoD da US-3.1 que, se o utilizador tiver >200 chats, a listagem é truncada silenciosamente — expor um `truncated`/`has_more` informativo na resposta (como o calendário faz com `truncated_at`, calendar.py L210-213) em vez de truncar sem sinal. Tratar na implementação.

### Notas

**A8 (Nota) — `lastMessagePreview` exige `$expand` próprio e/ou header `Prefer`.** §2.1 já o documenta e a tolerância a `None` está na DoD (R6). O `$expand=members` e `$expand=lastMessagePreview` podem precisar de ser combinados (`$expand=members,lastMessagePreview`); validar no tenant real. Conteúdo informativo já no plano.

**A9 (Nota) — FakeGraphClient: cobrir o `get_chat` se A2 for aceite.** §6.1 lista os fakes; se acrescentar `get_chat` (A2), estender o `__init__` com `chat` e o método fake correspondente, e a DoD da US-3.3 deve provar que o prepare chama `get_chat` (read) e **não** qualquer escrita (count `send_chat_message`=0).

**A10 (Nota) — `$orderby` em `/me/chats` pode não ser suportado.** §2.1 já prevê "ordena a tool" como fallback e a US-3.1 ordena defensivamente por `last_updated desc`. Correto. Igual cuidado vale para `/messages?$orderby=createdDateTime desc` (este é geralmente suportado). Sem ação.

---

## Parecer sobre as decisões D1–D11

- **D1 (criar 1:1 = escrita confirmada):** Sã. Tratar a criação como escrita (mesmo sendo idempotente no Graph) é a escolha conservadora e correta — "abrir conversa" é um efeito visível para terceiros. OK.
- **D2 (filtro client-side):** Sã e simples; o `$filter`/`$search` nos chats é mesmo inconsistente no Graph. OK. Acrescentar sinal de truncagem (A7).
- **D3 (procurar→criar):** Sã; ver A5 (fallback quando membro sem email) — aceite com documentação/teste.
- **D4 (default 25 / teto 50):** Sã, coerente com o email. OK.
- **D5 (N + `has_more`, não auto-pagina):** **Boa decisão** e bem justificada — diverge conscientemente do email (pergunta) e do calendário (auto-pagina) porque o histórico de chat pode ser enorme. A divergência é deliberada e documentada; aprovo.
- **D6 (default text; html a pedido):** Sã; validação de `body_type ∈ {text,html}` no prepare. OK.
- **D7 (sem @menções / sem `reply_to`):** Sã; reduz superfície. "Responder" = enviar no mesmo `chat_id` partilhando o par prepare/confirm — modelação correta (não há `reply` server-side em chats). OK.
- **D8 (`is_system`, cartões só metadados):** Sã; fronteira de não-confiável bem traçada. A derivação `messageType != "message"` é razoável (engloba `systemEventMessage` e outros). OK; cobrir no unit (§6.3).
- **D9 (resolve a montante):** Sã; idêntico ao Calendário. OK.
- **D10 (`_MAX_BODY_CHARS=28000`, erro orientador):** Sã; validado no prepare antes de qualquer token. OK.
- **D11 (sem reações/editar/eliminar):** Sã; diferimento coerente com v1.0/v1.1. OK.
- **Micro-decisão "auditoria `teams.send` usa `subject_hash(chat_id)`":** **Rejeitada — ver A1.** Deve ser removida.
- **Micro-decisão "listagem sem paginação para o cliente":** Aceite com a ressalva A7 (sinalizar truncagem).

---

## Pré-condições para o developer arrancar

**O developer PODE começar já** pela ordem §7.1 passos 1–3 (scopes, mapeadores/métodos Graph de leitura, US-3.1, US-3.2) — nada nesses passos depende dos achados Maiores.

**Antes de fechar as escritas (US-3.3 e US-3.4), corrigir no plano/contrato:**

1. **[A1 — Maior]** Remover o `subject_hash(chat_id)` do `extra` de `teams.send`. `extra` = `{chat_type, body_type}`; o `subject_hash` da identidade vem do `log_audit`; o `target=chat_id` já identifica o recurso. Atualizar §5.8 e §8.
2. **[A2 — Maior]** Acrescentar `get_chat(chat_id)` (read, `GET /me/chats/{id}?$expand=members`) e usá-lo no prepare do envio para montar o resumo (em vez do match sobre `list_chats(top=50)`). Estender o FakeGraphClient e a DoD (A9). Se rejeitado, assumir explicitamente o resumo degradado e reforçar a confirmação do chat no `instructions`/playbook.

**A tratar durante a implementação (não bloqueiam o arranque):**

3. **[A3]** Validar no tenant real o POST sob `/me/chats/.../messages`; registar como passo do runbook.
4. **[A4]** Simplificar `_chat_from` e unificar o mapeamento de membro; cobrir no unit.
5. **[A5]** Documentar e testar o caso "membro sem email → cria (idempotente)".
6. **[A7]** Sinalizar truncagem da listagem (`truncated`/`has_more`) em vez de truncar em silêncio.
7. **[A6/A8/A10]** Notas — sem ação obrigatória; refletir A6 em §7.2 (R4) para clareza.

---

### Resumo executivo para decisão

- **Veredicto:** APROVADO COM RESSALVAS.
- **Bloqueadores:** nenhum.
- **Maiores (corrigir no plano antes das escritas US-3.3/3.4):** A1 (colisão do `subject_hash` na auditoria — risco de privacidade) e A2 (resumo do envio degrada com >50 chats — enfraquece a confirmação humana).
- **Arranque imediato autorizado** para scopes + client + leituras (US-3.1/US-3.2). As escritas só depois de A1 e A2 fechados.
