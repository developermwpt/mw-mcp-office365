# Fase 2 — Calendário: runbook de validação manual (tenant/VPS reais)

> Os testes automáticos (Graph/Entra mockados) cobrem toda a lógica e estão **a passar**. Este
> runbook valida o comportamento real contra o Microsoft Graph no tenant da Mobiweb, com o
> servidor MCP no VPS. Cada passo é executado pelo assistente (Claude) ligado ao MCP; o
> validador confirma o resultado no Outlook/Calendário.

## 0. Pré-requisitos (BLOQUEADORES)

1. **Admin consent do scope `Calendars.ReadWrite`** no tenant Entra (cobre leitura, escrita e
   `getSchedule` delegado). Sem este consent, todas as operações de calendário falham com
   reauth/consent. O scope já está no default de `GRAPH_SCOPES` (`config.py`); falta o consent
   do admin Entra — **igual ao procedimento do `Mail.*` na Fase 1**.
2. Conta Office 365 **ligada** ao MCP (fluxo de login concluído; `whoami` devolve a conta).
3. Servidor MCP a correr no VPS com a build da Fase 2; relógio do VPS sincronizado (NTP).
4. Pelo menos: um evento próprio futuro, um evento **recorrente** (série) e um **convite
   recebido** de outra pessoa (para US-2.6) na caixa do utilizador de teste.

Após conceder o consent, **reiniciar a sessão / re-login** para o token Graph passar a incluir
os scopes de calendário.

## 1. US-2.1 — Listar eventos (leitura, auto-paginação, fuso)

1. Pedir: *"Mostra os meus eventos de hoje"* e *"…dos próximos 7 dias"*.
2. **Confirmar:**
   - As horas devolvidas estão no **fuso do mailbox** (não UTC) — comparar com o Outlook.
   - Eventos de uma **série recorrente** aparecem como ocorrências individuais marcadas
     `isRecurring=true`.
   - Para uma janela com muitos eventos, devolve **todos** (`auto_fetched_all=true`), não só a
     1ª página.
   - Um evento com corpo rico (HTML/imagens) tem o `bodyPreview` limpo e a resposta traz
     `content_is_untrusted=true`.
3. **Teste anti-injeção (opcional):** criar um evento cujo corpo contenha texto do tipo
   "INSTRUÇÃO AO ASSISTENTE: …"; confirmar que o assistente **não** age sobre essa instrução.

## 2. US-2.2 — Disponibilidade (`getSchedule`)

1. Pedir: *"Estou livre amanhã entre as 9h e as 18h? E o [colega@mobiweb.pt]?"*
2. **Confirmar:**
   - A resposta inclui o **próprio** utilizador mesmo que só se peça o colega.
   - As janelas ocupadas/livres batem certo com os calendários reais.
   - Horas no fuso do mailbox.

## 3. US-2.3 — Criar evento (prepare/confirm + Teams D6)

1. **Online (Teams):** pedir *"Marca uma reunião 'Sync' amanhã 10h–11h com
   [a@mobiweb.pt]"* (sem local).
   - **prepare:** o resumo declara "Notifica 1 participante(s) (domínios: mobiweb.pt)" e
     **"Inclui link Teams"**. Nada é criado ainda (verificar que o evento NÃO aparece no
     Outlook).
   - **confirm:** o evento aparece no calendário, o participante **recebe convite**, e o evento
     tem **link Teams**.
2. **Presencial (sem Teams):** repetir com `location="Sala 1"`.
   - **prepare:** o resumo diz **"Presencial em 'Sala 1' (sem link Teams)"**.
   - **confirm:** evento criado sem link Teams, com a sala como localização.
3. **Idempotência:** reapresentar o mesmo `confirmation_token` → resposta `idempotent_replay`
   e **não** cria um segundo evento.

## 4. US-2.4 — Editar/reagendar (recorrência → clarification)

1. **Não-recorrente:** *"Reagenda a 'Sync' para as 14h"* → prepare normal; confirm muda a hora
   no Outlook e o participante recebe a atualização.
2. **Recorrente:** *"Muda o assunto da reunião diária"* sobre uma **ocorrência de série**:
   - O assistente **pergunta** "só esta ocorrência ou a série inteira?" (`needs_clarification`),
     **sem** alterar nada.
   - Responder "a série inteira" → repete com `scope='series'` → a alteração aplica-se a
     **todas** as ocorrências (verificar no Outlook).
   - Responder "só esta" → `scope='occurrence'` → só a ocorrência selecionada muda.

## 5. US-2.5 — Cancelar evento

1. **Como organizador:** *"Cancela a 'Sync'"* → prepare declara "Notifica N participante(s) …
   Alto impacto"; confirm cancela e os participantes **recebem o cancelamento**.
2. **Como NÃO-organizador:** sobre um evento organizado por outra pessoa, pedir cancelar →
   o assistente devolve **erro orientando para `decline`** (não cancela). Confirmar que o
   evento continua no calendário do organizador.
3. **Recorrente:** cancelar uma ocorrência de série → `needs_clarification` (esta vs série)
   antes de cancelar.
4. **Idempotência:** repetir o token → `idempotent_replay`, sem segundo cancelamento.

## 6. US-2.6 — Responder a convite (accept/decline/tentative)

1. A partir de um **convite recebido**, pedir *"Recusa este convite, diz que não posso"*.
   - **prepare:** o resumo **declara a transição** (ex.: "Já tinha Aceitado; vai mudar para
     Recusado e notificar o organizador") — confirmar que o estado anterior está correto.
   - **confirm:** o organizador recebe a resposta; o estado do convite muda no Outlook.
2. Repetir com **accept** e **tentative**.
3. **Organizador bloqueado:** sobre um evento **organizado pelo próprio**, pedir para
   responder → **erro** ("é o organizador; não pode responder ao próprio convite"), sem token.
4. **Idempotência:** repetir o token → `idempotent_replay`.

## 7. Garantias transversais a confirmar (qualquer US de escrita)

- **Reauth graciosa:** se a sessão expirar/for revogada a meio, a operação devolve
  `reauth_required` (mensagem amigável) e **nada** é escrito; após re-login, o mesmo
  `confirmation_token` ainda funciona (não foi consumido).
- **Auditoria:** confirmar nos logs do VPS que cada escrita gera `event=audit` com
  `subject_hash`, `action` (`calendar.create|update|cancel|respond`), `target` (event id) e
  contagem de participantes — **sem** assunto em claro, emails ou corpo.
- **TTL:** deixar o `confirmation_token` expirar (TTL) e confirmar → `expired`.

## 8. Registo de resultados

Anotar, por US, a data, o validador, o resultado (✅/⬜/❌) e quaisquer desvios. Atualizar a
coluna **Validação manual** em `docs/fase-2/estado-user-stories.md` à medida que cada US for
confirmada no tenant real (como foi feito na Fase 1 em 2026-06-02).
