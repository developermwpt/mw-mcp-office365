# Análise Funcional — Módulo de Aprendizagem de Comportamento de Email
## Fase Aprendizagem (US-L.x)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de Office 365 — módulo de aprendizagem do comportamento do utilizador no email
**Data:** 2026-06-02
**Estado:** Análise funcional + scaffold funcional e testado. **Não liga ao Microsoft Graph em produção** além do que já existe — a aprendizagem trabalha sobre metadados locais já disponíveis e sobre um novo store local.

> **Relação com a [v1.1](../analise-funcional-v1.1.md):** este módulo é uma fase adicional, ortogonal aos módulos funcionais (Email, Calendário, Teams, Ficheiros). Reutiliza integralmente a arquitetura de identidade dual-plane (§2 da v1.1), o modelo de aprovação em duas fases (§3) e a auditoria só-metadados (§1.2/§7). Como a Fase 2 candidata é "Calendário" (US-2.x), usa-se o prefixo **US-L.x** (*Learning*) para não colidir.

---

## 1. Objetivo e princípios

Aprender, a partir das **ações que o utilizador confirma** sobre os seus emails (mover, arquivar, responder, reencaminhar, eliminar, enviar), padrões reutilizáveis e, perante um novo email "parecido", **sugerir** a ação habitual — que o utilizador apenas tem de **confirmar**. A sugestão nunca é executada automaticamente: a execução passa SEMPRE pelo fluxo `*_prepare` → `confirmation_token` → `*_confirm` já existente.

Princípios (herdados e reforçados da v1.1):
- **Opt-in explícito** — a aprendizagem está desligada por defeito; nada é registado sem consentimento.
- **Só-metadados** — nunca o corpo dos emails nem PII em claro.
- **Isolamento estrito por `subject`** — sem acesso cruzado entre utilizadores.
- **Explicável e auditável** — cada recomendação traz um `rationale` em linguagem natural.
- **Anti prompt injection** — as features derivam de metadados, não de instruções do corpo.
- **Nunca auto-executar** — recomendar ≠ executar.
- **Degradação graciosa** — opt-out, sem histórico ou erro de store → resposta amigável; o registo de comportamento NUNCA quebra a operação de email principal.

## 2. Modelo de privacidade (decisão justificada)

> **Decisão:** **opt-in explícito + só-metadados + isolamento por `subject` + retenção limitada + sem corpo/PII em claro + cifra do payload sensível.** É a abordagem mais conservadora e a única coerente com o risco CNPD/RGPD já registado na v1.1 (§7) e com a auditoria só-metadados (`observability/audit.py`).

**Porquê esta escolha (trade-offs ponderados):**

| Alternativa considerada | Porque foi rejeitada |
|---|---|
| Guardar corpo/assunto em claro para "melhor aprendizagem" | Multiplicaria o impacto de um comprometimento da VPS (já um risco aceite com mitigações na v1.1 §4) e tornaria a base legal RGPD muito mais frágil. O ganho de qualidade não compensa: o domínio do remetente + tokens do assunto chegam para padrões úteis. |
| Aprendizagem ligada por defeito (opt-out) | Sem consentimento explícito não há base legal limpa para tratamento de novos dados comportamentais. Opt-in é defensável perante a CNPD. |
| Modelo estatístico/ML opaco | Não-explicável e não-auditável; impossível dar um `rationale` honesto ao utilizador. Uma heurística linear é suficiente e transparente. |

**Concretização:**
- **Consentimento:** `learning_opt_in(enabled=True/False)`. Default desligado (a nível de servidor via `LEARNING_ENABLED=false` e por utilizador via `learning_preferences.opt_in=0`). Sem opt-in, `record_action_event` não regista e `email_recommendations` devolve `status=opt_out` com instruções de ativação.
- **Ativar/desativar:** a qualquer momento, pela mesma tool. Desativar pára o registo de novos eventos; o histórico mantém-se até ser apagado.
- **Apagar dados (direito ao esquecimento):** `learning_forget` apaga TODO o histórico do `subject`.
- **Retenção limitada:** `LEARNING_RETENTION_DAYS` (default 180); `purge_behavior_events(subject, before=...)` permite purga de eventos antigos (a agendar por operações, à imagem da retenção de logs).
- **Cifra:** a assinatura de features (que pode conter tokens de assunto) é serializada em JSON e **cifrada** com `Cipher` antes do disco (`behavior_events.features_enc`), como os tokens Graph. A `action`, o `sender_domain` e a pasta `destination` ficam em claro para indexação leve — são metadados de baixo risco (o domínio já aparece, no máximo, na auditoria da v1.1).
- **Auditoria sem PII:** eventos `learning.event_recorded`, `learning.recommend`, `learning.opt_in`, `learning.forget` registam apenas `subject_hash`, contagens e, no máximo, o domínio.

## 3. Modelo de dados de comportamento

**Sinais registados (só metadados):**
- **Ação tomada:** `move` | `archive` | `reply` | `reply_all` | `forward` | `delete` | `send`.
- **Remetente:** apenas o **domínio** (`sender_domain`), nunca a mailbox local.
- **Assunto:** reduzido a **tokens normalizados** (minúsculas, sem pontuação, sem prefixos de resposta, ≥3 chars, únicos e ordenados) — suficiente para similaridade, insuficiente para reconstruir o texto.
- **Flags:** `has_attachments`, `is_reply`, `importance` (low/normal/high), `is_newsletter` (presença de `List-Id`/`List-Unsubscribe`).
- **Pasta:** `destination` (origem/destino do move), em claro.
- **Nunca:** o corpo, endereços completos, anexos.

**Nova tabela `behavior_events`** (ver `storage/schema.sql`):

```sql
CREATE TABLE behavior_events (
    event_id      TEXT PRIMARY KEY,
    subject       TEXT NOT NULL,           -- isolamento
    action        TEXT NOT NULL,           -- move|archive|reply|reply_all|forward|delete|send
    sender_domain TEXT,                    -- domínio (sem mailbox local)
    destination   TEXT,                    -- pasta destino quando aplicável
    features_enc  BLOB NOT NULL,           -- JSON cifrado da assinatura (Cipher)
    created_at    TEXT NOT NULL            -- ISO 8601 UTC (retenção/purga)
);
-- índices: (subject) e (subject, sender_domain)
```

**Nova tabela `learning_preferences`** (opt-in por subject):

```sql
CREATE TABLE learning_preferences (
    subject     TEXT PRIMARY KEY,
    opt_in      INTEGER NOT NULL DEFAULT 0,  -- 0 desligado (default), 1 consentido
    updated_at  TEXT NOT NULL
);
```

Métodos no `TokenStore` (mesmo estilo: lock, clock, cifra, filtro por subject): `set_learning_opt_in`, `get_learning_opt_in`, `record_behavior_event`, `list_behavior_events`, `purge_behavior_events`.

## 4. Extração de features e similaridade

A **assinatura** de um email (`EmailSignature`, em `learning/features.py`) deriva apenas de metadados do dict Graph (`from`/`sender`, `subject`, `hasAttachments`, `importance`, cabeçalhos `List-*`). A função `extract_signature` é a **fronteira anti prompt injection**: o que entra é tratado como dado, nunca como instrução; o `body` é ignorado.

**Similaridade** (`similarity(a, b)` → `[0,1]`): combinação linear, deliberadamente transparente e auditável:

```
score =  0.60 · [mesmo sender_domain]
       + 0.25 · Jaccard(subject_tokens_a, subject_tokens_b)
       + 0.05 · [mesma flag has_attachments]
       + 0.05 · [mesma flag is_reply]
       + 0.05 · [mesma flag is_newsletter]
       (saturado em 1.0)
```

O domínio do remetente é o sinal dominante e estável; os tokens do assunto afinam dentro do mesmo domínio. **Porque é explicável:** cada parcela é inspecionável e mapeia diretamente para uma frase do `rationale` ("Costumas mover emails de @newsletter.acme.com para 'Archive' (8 vezes)").

**Score de recomendação** (`learning/recommender.py`): agrupam-se os eventos passados por `(ação, destino)`; para cada grupo:

```
confidence = média(similaridade do alvo com cada evento do grupo) · fator_de_suporte
fator_de_suporte = min(nº_eventos / 5, 1.0)     # satura aos 5 eventos
```

Só entram grupos com `nº_eventos ≥ 2` e `confidence ≥ LEARNING_MIN_CONFIDENCE`. Devolve no máximo `LEARNING_TOP_N`, ordenado por confiança (desempate determinístico).

## 5. Fluxo de recomendação (integrado no two-phase approval)

```
                 (read-only, opt-in)
email_read  ──►  email_recommendations(message, message_id)
                      │
                      ├─ opt-out? ──► status=opt_out + como ativar  (FIM, nada executado)
                      │
                      └─ sugestões ordenadas: { action, params, confidence, rationale,
                                                prepare_tool, prepare_params }
                                   │
   utilizador ACEITA uma sugestão  │  (decisão humana — nunca automático)
                                   ▼
            chamar o prepare_tool indicado  (ex.: email_move_prepare)
                                   │
                                   ▼
                         confirmation_token  ── TTL ──►  email_*_confirm(token)
                                                              │
                                                              ▼
                                                   executa no Graph + auditoria
                                                   + (se opt-in) record_action_event
```

- **Limiar de confiança:** `LEARNING_MIN_CONFIDENCE` (default 0.5); **top-N:** `LEARNING_TOP_N` (default 3).
- **Sem auto-execução:** `email_recommendations` é estritamente read-only e não devolve qualquer `confirmation_token`. O token só nasce quando o utilizador, por intenção direta, chama o `*_prepare`. Não há segundo mecanismo de confirmação.
- **Ciclo virtuoso:** cada `*_confirm` chama `record_action_event` (defensivo, só se opt-in), pelo que aceitar recomendações reforça o aprendizado.

## 6. User Stories (US-L.x)

| US | Título | Critérios de aceitação |
|----|--------|------------------------|
| **US-L.1** | Registar comportamento de email | Ao confirmar um move/reply/forward/delete/send com opt-in ligado, é gravado um `behavior_event` só com metadados (ação, domínio, destino, assinatura cifrada). Sem opt-in, nada é gravado. Uma falha do store NÃO quebra a operação de email. |
| **US-L.2** | Gerar recomendações para um email | Dado um email (metadados) com opt-in, `email_recommendations` devolve até `top_n` sugestões ordenadas por confiança, cada uma com `action`, `params`, `confidence`, `rationale` e o `prepare_tool`/`prepare_params`. Sem histórico relevante → lista vazia. Read-only, sem chamadas Graph. |
| **US-L.3** | Aceitar uma recomendação | Aceitar uma sugestão consiste em chamar o `prepare_tool` indicado com os `prepare_params`, que devolve um `confirmation_token` do fluxo existente; o `*_confirm` executa. Nenhuma ação é executada pela própria tool de recomendação. |
| **US-L.4** | Ignorar/não agir sobre recomendações | Recomendações não aceites não têm efeito nem persistem estado. (Feedback explícito de "ignorar" para ajuste fino fica em backlog — ver §11.) |
| **US-L.5** | Opt-in / opt-out | `learning_opt_in(enabled)` ativa/desativa por utilizador. Default desligado. Devolve mensagem clara do efeito. Desativar pára o registo; o histórico mantém-se. |
| **US-L.6** | Apagar histórico (esquecimento) | `learning_forget` apaga TODO o histórico do `subject` e devolve a contagem apagada. Isolado por subject. |

## 7. RGPD/Compliance, Riscos, Testes, Faseamento, Decisões em aberto

**RGPD/Compliance:** base legal = **consentimento explícito** (opt-in); **minimização** (só metadados, sem corpo/PII em claro); **limitação de conservação** (retenção `LEARNING_RETENTION_DAYS` + purga); **direito ao esquecimento** (`learning_forget`); pseudonimização nos logs (`subject_hash`). Reforça — não contraria — a postura da v1.1 §7. Recomenda-se incluir este tratamento na DPIA.

**Riscos e mitigações:**

| Risco | Mitigação |
|---|---|
| **Prompt injection** via conteúdo lido | As features derivam SÓ de metadados (`extract_signature` ignora o `body`); o conteúdo nunca é tratado como instrução. A recomendação não executa — só o utilizador, por intenção direta, aciona o `*_prepare`. |
| Auto-execução indevida | Impossível por construção: a tool de recomendação é read-only e não emite tokens. |
| Sobre-recomendação / fadiga | Limiar de confiança + suporte mínimo + top-N conservadores. |
| Recomendação errada após mudança de hábito | Score pondera suporte; `learning_forget` permite recomeçar; retenção limita peso de hábitos antigos. |
| Fuga de dados comportamentais | Features cifradas; isolamento por subject; só metadados. |

**Testes/observabilidade:** unitários (`test_learning_features.py`, `test_learning_recommender.py`, `test_learning_store.py`) e integração (`test_learning_e2e.py`), todos sem rede e com relógio controlável. Auditoria: eventos `learning.*` só-metadados.

**Faseamento:**
1. **Scaffold (esta entrega):** store, features, recommender, tools read-only, opt-in/forget, registo nos confirms de email, testes.
2. **Endurecimento:** purga agendada por retenção; métricas de aceitação de recomendações; afinação de pesos com dados reais.
3. **Extensão:** feedback explícito (US-L.4 avançada), recência no score, alargar a Calendário/Teams.

**Decisões em aberto (escalar ao cliente):**
- Confirmar **opt-in por utilizador** (decisão atual) vs autorização do administrador a nível de tenant.
- `LEARNING_RETENTION_DAYS` final e quem agenda a purga (ops vs job interno).
- Se e como expor **feedback explícito** ("não voltar a sugerir isto") — implica guardar sinais negativos.
- Incluir o tratamento comportamental na **DPIA** e no registo de atividades de tratamento.
