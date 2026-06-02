# Análise Funcional — Contactos e Resolução de Destinatários
## Módulo 5 (US-5.x)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de Office 365 — resolução de nome → destinatário
**Data:** 2026-06-02
**Estado:** Implementado e **validado no tenant real (2026-06-02)**. Read-only sobre o Graph.

> **Relação com a [v1.1](../analise-funcional-v1.1.md):** módulo transversal aos módulos
> funcionais (serve Email e, na Fase 2, o Calendário). Reutiliza a arquitetura de identidade
> dual-plane (§2), o padrão de **confirmação** (o mesmo `needs_clarification` usado em
> `email_reply_prepare`) e a auditoria só-metadados (§1.2). Read-only: a resolução nunca
> envia nem agenda nada — apenas devolve candidatos para o utilizador escolher.

---

## 1. Problema e objetivo

Hoje as operações de escrita exigem o **endereço completo** (`to=["vera@..."]`). Quando o
utilizador diz *"manda o email à Vera"* ou *"marca a reunião com a Vera"*, o assistente não
tem forma fiável de saber **quem é "a Vera"** — e adivinhar um endereço é inaceitável (pode
enviar para a pessoa errada).

**Objetivo:** dado um **nome** (ou nome parcial), procurar nos contactos/pessoas do utilizador
e devolver **candidatos ordenados por relevância** (nome + email), para o utilizador
**confirmar** qual antes de qualquer ação. A resolução é **leitura**; o envio/agendamento
continua a passar pelo `*_prepare` → `confirmation_token` → `*_confirm`.

## 2. Fontes de dados (Microsoft Graph, delegated, read-only)

| Fonte | Endpoint | Scope | Papel |
|---|---|---|---|
| **People** (primária) | `GET /me/people?$search="vera"` | `People.Read` | Ranqueia pelas pessoas com quem o utilizador mais comunica (colegas + contactos) — o sinal mais útil |
| **Contactos pessoais** | `GET /me/contacts?$search="vera"` | `Contacts.Read` | Agenda pessoal do utilizador (fallback / complemento) |
| Diretório da organização | `GET /users?$search="displayName:vera"` | `User.Read.All` (admin) | **Fora de âmbito inicial** — scope mais sensível; só se necessário |

Começa-se com **People + Contactos**. O diretório completo (`User.Read.All`) fica diferido
por ser um scope largo (lê todo o diretório) — decisão de risco a tomar com o cliente.

## 3. User Stories

| US | Título | Critérios de aceitação |
|----|--------|------------------------|
| **US-5.1** | Resolver destinatário por nome | `resolve_recipient(name)` devolve candidatos `{display_name, email, source}` ordenados por relevância. **0 resultados** → `not_found` com mensagem. **1** → `status=ok` com o candidato (o assistente propõe; o envio confirma). **Vários** → `needs_clarification` para o utilizador escolher. Read-only, sem `confirmation_token`. |
| **US-5.2** | Confirmar antes de usar | Nenhuma ação (enviar/agendar) usa um email resolvido sem confirmação humana: o assistente apresenta o candidato e só depois chama o `*_prepare` com o endereço escolhido. Quando ambíguo, **nunca** escolhe sozinho. |

## 4. Fluxo (integrado)

```
"manda à Vera"  ─►  resolve_recipient("vera")   (read-only)
                        │
                        ├─ 0 ─► not_found ("não encontrei 'vera' nos contactos")   (FIM)
                        ├─ 1 ─► ok + candidato  ─►  o assistente propõe "vera@x.pt?"
                        └─ N ─► needs_clarification + lista  ─►  utilizador escolhe
                                                                     │
                                       (email escolhido, por intenção do utilizador)
                                                                     ▼
                                       email_send_prepare(to=[email], ...) ─► confirm
```

## 5. Segurança e privacidade

- **Anti prompt injection:** a pesquisa é despoletada pelo **nome dado pelo utilizador**,
  não por conteúdo de emails; os candidatos são sempre confirmados por humano.
- **Só leitura:** `resolve_recipient` não emite `confirmation_token` nem executa nada.
- **Isolamento por `subject`:** usa o token Graph do próprio utilizador (vê só os seus
  contactos/people).
- **Auditoria só-metadados:** evento `contacts.resolve` com `subject_hash`, o termo de
  pesquisa (curto, dado pelo utilizador) e a contagem de candidatos — nunca a lista de
  endereços.
- **RGPD:** lê dados pessoais de terceiros (nomes/emails de colegas). Acresce ao tratamento
  a registar na **DPIA pendente** (ver tarefa de RGPD). Minimização: devolve só nome+email
  dos candidatos relevantes, não exporta a agenda inteira.

## 6. Scopes a adicionar (pré-requisito de produção)

`People.Read` e `Contacts.Read` (delegated) — adicionar no registo da app Entra + **admin
consent** + atualizar `GRAPH_SCOPES` na VPS (como se fez para `Mail.*`). Até lá, o módulo
funciona em testes (Graph mockado) mas as chamadas reais devolverão falta de consentimento.

## 7. Faseamento e decisões em aberto

1. **Esta entrega:** `resolve_recipient` (People + Contactos), padrão de confirmação, testes.
2. **Extensão:** ligar ao Calendário (Fase 2) para *"marca reunião com a Vera"*; afinar o
   ranking com o histórico do módulo de aprendizagem (a quem o utilizador costuma escrever).
3. **Decisões em aberto (cliente):** incluir ou não o **diretório da organização**
   (`User.Read.All`, scope largo); incluir este tratamento na **DPIA**.
