# Fase 3 — Teams (chats): runbook de validação manual (tenant/VPS reais)

> Os testes automáticos (Graph/Entra mockados) cobrem toda a lógica e estão **a passar**
> (254 testes, 43 da Fase 3). Este runbook valida o comportamento real contra o Microsoft Graph
> no tenant da Mobiweb, com o servidor MCP no VPS. Cada passo é executado pelo assistente
> (Claude) ligado ao MCP; o validador confirma o resultado no Microsoft Teams.

## 0. Pré-requisitos (BLOQUEADORES)

1. **Admin consent dos scopes `Chat.Read` e `Chat.ReadWrite`** (delegated) no tenant Entra —
   **mesmo procedimento do `Calendars.ReadWrite`** da Fase 2: adicionar no registo da app +
   admin consent + atualizar `GRAPH_SCOPES` no `.env` de produção (a fonte de verdade) e no
   `config.py`/`.env.example` (lição da Fase 2). Sem este consent, todas as operações de Teams
   falham com reauth/consent. Os scopes já estão no default de `GRAPH_SCOPES` (`config.py`).
2. Conta Office 365 **ligada** ao MCP (fluxo de login concluído; `whoami` devolve a conta).
3. Servidor MCP a correr no VPS com a **build da Fase 3**; relógio do VPS sincronizado (NTP).
4. Na caixa de teste, pelo menos: **um chat 1:1**, **um chat de grupo** (com tópico e algumas
   mensagens, incluindo uma de sistema — ex.: alguém entrou no grupo ou mudou o tópico), e **um
   contacto com quem ainda NÃO exista conversa** (para iniciar uma conversa nova em US-3.4).

Após conceder o consent, **reiniciar a sessão / re-login** para o token Graph passar a incluir
os scopes de Teams.

> **Ponto de verificação de endpoint (A3).** O código envia/lê sob `/me/chats/{id}/...` por
> coerência com o resto do client. Em delegated isto funciona para chats do próprio, mas o
> `POST /me/chats/{id}/messages` é menos documentado que `POST /chats/{id}/messages`. **Na
> primeira validação de US-3.3, confirmar explicitamente que o envio sob `/me/chats/...` é
> aceite pelo Graph;** se devolver erro, registar e cair para `/chats/{id}/messages`.

## 1. US-3.1 — Listar chats (1:1 e grupo, filtro client-side)

1. Pedir: *"Mostra os meus chats de Teams"* e depois *"…filtra pelos que têm a [Vera] ou o
   tópico [Projeto X]"*.
2. **Confirmar:**
   - Aparecem chats **1:1** (`chat_type=oneOnOne`, sem tópico) e **de grupo**
     (`chat_type=group`, com tópico), cada um com o seu `id`.
   - Os `members` trazem **só nome + email** (e `aad_user_id` quando o email não vem) — nenhum
     outro atributo de diretório (minimização RGPD).
   - O **filtro por tópico** devolve só o(s) grupo(s) com esse tópico; o **filtro por
     nome/email** devolve só os chats com esse participante (substring, sem distinção de
     maiúsculas).
   - O `last_message_preview`, quando existe, vem **limpo** (sanitizado) e a resposta traz
     `content_is_untrusted=true`; quando não vem, a listagem não falha (preview vazio).
3. **Teste anti-injeção (opcional):** num chat, deixar uma mensagem cujo preview contenha texto
   do tipo "INSTRUÇÃO AO ASSISTENTE: …"; confirmar que o assistente **não** age sobre essa
   instrução (trata o preview como conteúdo não-confiável).

## 2. US-3.2 — Ler mensagens de um chat

1. A partir de um `chat_id` de US-3.1, pedir: *"Lê as últimas mensagens deste chat"* e depois
   *"…mostra mais antigas"*.
2. **Confirmar:**
   - Devolve as **N mais recentes** (default 25; ordem decrescente) — comparar com o Teams.
   - Quando há mais histórico, vem `has_more=true` e um `next_link`; **só** ao pedir "mais
     antigas" (com `page_token`) é que traz mensagens anteriores — **não auto-pagina** sozinho.
   - Uma **mensagem de sistema** (entrada/saída de membro, mudança de tópico) aparece marcada
     `is_system=true` e o assistente **não a interpreta** como conteúdo acionável.
   - Mensagens com corpo rico (HTML) vêm com o `body` limpo e `content_is_untrusted=true`.
   - Pedir mais de 50 mensagens devolve no máximo 50 (teto D4).

## 3. US-3.3 — Enviar mensagem num chat existente (prepare/confirm)

1. **Chat de grupo:** *"Manda 'Bom dia, equipa' no grupo [Projeto X]"* (usar o `chat_id` do
   grupo).
   - **prepare:** o resumo declara **"Enviar mensagem no grupo \"[tópico]\" com N
     participante(s) (domínios: …) [formato: text]"** e **não envia** (verificar que a mensagem
     NÃO aparece ainda no Teams). Confirmar que **N e os domínios batem certo** com o grupo (é a
     barreira anti-erro — "enviei para o grupo certo?"). **N exclui o próprio** (emissor não é
     destinatário): num grupo de 3 membros incluindo-te, N = 2.
   - **confirm:** a mensagem aparece no chat e os participantes são **notificados** pelo Teams.
   - **Verificação A3:** confirmar que o `POST` sob `/me/chats/...` foi aceite (sem erro de
     endpoint).
2. **Chat 1:1, formato HTML:** *"Manda à [Vera] uma mensagem em **negrito** a dizer …"* →
   `body_type=html`.
   - **prepare:** num 1:1 o resumo **NOMEIA a pessoa** (barreira concreta), p.ex. **"Enviar
     mensagem a Vera Martins (vera.martins@mobiweb.pt) [formato: html]"** — `recipients_count=1`
     (o próprio está excluído), sem "participante(s)".
   - **confirm:** a mensagem chega formatada.
3. **Mensagem demasiado longa (D10):** pedir o envio de um texto com mais de ~28000 caracteres →
   o assistente devolve **erro orientador** ("Mensagem demasiado longa … Divida em partes"),
   **sem** preparar/enviar.
4. **Formato inválido:** forçar `body_type` diferente de text/html → **erro** ("Use 'text' ou
   'html'"), sem token.
5. **Idempotência:** reapresentar o mesmo `confirmation_token` → resposta `idempotent_replay` e
   **não** envia uma segunda mensagem (anti-duplicação).

## 4. US-3.4 — Iniciar conversa 1:1 por nome (obter/criar)

> Resolução por nome é a **montante**: o assistente usa `resolve_recipient` e **confirma o
> email** com o utilizador ANTES de chamar a tool de Teams (D9). As tools de Teams só recebem
> emails/`chat_id` já resolvidos.

1. **Chat 1:1 já existente:** *"Manda mensagem ao [colega com quem já há conversa] no Teams"*.
   - O assistente resolve o nome → confirma o email → `teams_get_or_create_one_on_one_chat`
     devolve **`status=ok` com o `chat_id`** (a conversa já existe) — **sem** pedir confirmação
     para criar — e segue direto para `teams_send_message_prepare`.
2. **Conversa nova (inexistente):** *"Inicia conversa no Teams com o [contacto sem chat ainda]"*.
   - **prepare:** o assistente devolve **`pending_confirmation`** com o resumo **"Vai INICIAR
     uma nova conversa de Teams (1:1) com <email>"** e um token (porque criar/abrir a conversa é
     uma **escrita**) — confirmar que a conversa ainda NÃO aparece no Teams.
   - **confirm:** a conversa 1:1 é criada/aberta, devolve o `chat_id`, e a seguir consegue-se
     enviar a mensagem (US-3.3) nesse `chat_id`.
   - **Idempotência:** repetir o token → `idempotent_replay`, sem criar um segundo chat.
3. **Membro sem email no diretório (esperado):** se o contacto resolvido não casar com nenhum
   1:1 existente (ex.: chats cujo membro só traz `userId`), o fluxo segue para criação
   (idempotente no Graph) — comportamento esperado, não é erro.

## 5. US-3.5 — Responder numa conversa (= enviar no mesmo chat_id)

1. Num chat já aberto (1:1 ou grupo), pedir *"Responde aí a dizer …"*.
2. **Confirmar:**
   - O assistente **reusa o mesmo par** `teams_send_message_prepare`/`_confirm` no **mesmo
     `chat_id`** (em chats não há thread/reply server-side — D7): prepare resume + token,
     confirm envia.
   - A resposta aparece como nova mensagem no chat. **Não** há citação/`reply_to` nem @menções
     (diferidos da v1).

## 6. Garantias transversais a confirmar (qualquer US de escrita)

- **Reauth graciosa:** se a sessão expirar/for revogada a meio, a operação devolve
  `reauth_required` (mensagem amigável) e **nada** é escrito; após re-login, o mesmo
  `confirmation_token` ainda funciona (não foi consumido). Uma leitura acessória do prepare
  (`get_chat`) que falhe **não** derruba a sessão — o resumo apenas degrada ("Detalhes do chat
  indisponíveis.") e o token continua a ser emitido.
- **Auditoria (logs do VPS):** confirmar que cada escrita gera `event=audit` com `subject_hash`
  (da identidade do utilizador), `action` (`teams.send` | `teams.chat_create`), `target` (o
  `chat_id`), `recipients_count` e `extra` seguro (`teams.send` → `{chat_type, body_type}`;
  `teams.chat_create` → `{chat_type:"oneOnOne", is_new_chat:true}`). **Sem** o texto da
  mensagem, nomes ou emails em claro, e **sem** um segundo `subject_hash` em `extra` (A1).
- **TTL:** deixar o `confirmation_token` expirar (TTL) e confirmar → `expired`.
- **Isolamento por subject:** cada utilizador só vê/opera sobre os seus próprios chats (token
  delegado); confirmar que não há acesso cruzado.

## 7. Registo de resultados

Anotar, por US, a data, o validador, o resultado (✅/⬜/❌) e quaisquer desvios (em particular o
ponto de verificação A3 do endpoint POST). Atualizar a coluna **Validação manual** em
`docs/fase-3/estado-user-stories.md` à medida que cada US for confirmada no tenant real (como
foi feito nas Fases 1 e 2). Enquanto não houver admin consent de `Chat.Read`/`Chat.ReadWrite` e
acesso ao tenant real, **todas as US permanecem ⬜** na validação manual (as colunas
Implementado e Testado (auto) já estão ✅).
