# Runbook de Validação Manual — PoC Fase 0

**Projeto:** mw-mcp-office365 (Mobiweb)
**Executado por:** cliente (admin Entra + acesso à VPS), no **tenant e VPS reais**.
**Objetivo:** gerar a evidência de go/no-go que o código + testes automáticos **não podem**
produzir — o comportamento real do connector no Claude Team e do **refresh sob a Conditional
Access** do tenant (ver [plano §1.2/1.3](plano-implementacao.md) e [v1.1 §2.4](../analise-funcional-v1.1.md)).

> **Porque é manual.** Os testes automáticos provam que o *código* trata o dual-plane, o
> refresh e o `invalid_grant` (com Entra/Graph mockados). **Não** provam o que o Entra real
> faz sob a política de CA, nem se o connector Claude Team liga (esp. Mobile). Isso só se
> observa neste runbook.

---

## Pré-requisitos

- [ ] Acesso **admin** ao tenant Entra ID; **export da política de Conditional Access** atual.
- [ ] VPS com `mw-mcp-office365.mobiweb.pt` atrás de Cloudflare (**Origin Certificate, Full
      strict**); porta de origem a servir o MCP.
- [ ] Python 3.12 na VPS; o servidor instalado de forma **não-editável**
      (`pip install .` — evita o caveat do `.pth` com caminhos com espaços; ver
      [src/README.md](../../src/README.md)).
- [ ] 1 utilizador de teste num **dispositivo gerido/compliant**.
- [ ] Claude **Team** com **Desktop** e **Mobile** disponíveis.
- [ ] Acesso aos **logs estruturados** do servidor (stdout JSON) — em especial o evento
      `refresh_failure`.

---

## Passo 0 — Registo da app no Entra ID

1. Registar uma aplicação **client confidencial**.
2. **Redirect URI** = `https://mw-mcp-office365.mobiweb.pt/callback` (o `OAUTH_REDIRECT_URI`).
3. Gerar **client secret**.
4. Adicionar os **scopes delegados** da PoC: `User.Read`, `offline_access`, `openid`, `profile`.
5. Conceder **admin consent global** (tenant-wide) desses scopes.

> **Critério:** após o consent global, o login do utilizador **não** deve mostrar ecrã de
> consentimento individual. Se mostrar → ver risco R5 (ajuste de config, não fatal).

## Passo 1 — Configurar e arrancar o servidor

1. Preencher o `.env` na VPS a partir de [`.env.example`](../../.env.example) (tenant, client
   id/secret, authority, redirect, chave de cifra base64 de 32 bytes, domínio).
2. Arrancar o servidor; garantir que escuta atrás do Cloudflare.
3. **Verificar liveness/readiness:**
   - `GET https://mw-mcp-office365.mobiweb.pt/healthz` → `200 {"status":"ok"}`
   - `GET .../readyz` → `200 {"status":"ready","db":"ok"}`

## Passo 2 — Validar a metadata pública (TLS no domínio final)

- `GET /.well-known/oauth-protected-resource` → 200, JSON com `resource` e
  `authorization_servers` (RFC 9728).
- `GET /.well-known/oauth-authorization-server` → 200, JSON com `issuer`,
  `authorization_endpoint`, `token_endpoint`, `registration_endpoint` (RFC 8414 + 7591).

> **Critério:** ambos respondem com TLS válido pelo **domínio final**. Falha aqui → R4
> (metadata/DCR incompatíveis) — ajustar antes de prosseguir.

## Passo 3 — Ligar o connector no Claude Team **Desktop**  → **(G1, G2)**

1. Adicionar o connector remoto apontando para `https://mw-mcp-office365.mobiweb.pt/mcp`.
2. Observar: descoberta de metadata → **DCR** (registo automático) → **login interativo**
   (browser, no dispositivo gerido) → conclusão do `/callback`.
3. Executar a tool **`whoami`**.

> **Critério (G2):** `whoami` devolve `displayName`/`userPrincipalName` corretos → o
> **dual-plane fecha end-to-end**. Registar o `home_account_id` para o passo 5.

## Passo 4 — Ligar o connector no Claude Team **Mobile**  → **(G1 / N1)**

1. Repetir a ligação e o login no **Mobile**.
2. Executar `whoami`.

> **Ponto de maior risco de plataforma.** Se o connector **não ligar** no Mobile (limitação
> da plataforma Claude) → **N1** (no-go parcial para Mobile). Documentar o comportamento exato
> (onde falha: descoberta, DCR, login, ou chamada).

## Passo 5 — Observar o **refresh sob Conditional Access**  → **(G3 / N2)** ⚠️ ponto central

1. Forçar a expiração do access token Graph: **reduzir a lifetime** do token no Entra (token
   lifetime policy) **ou** aguardar a expiração natural (~1h).
2. Voltar a executar `whoami` (obriga o servidor a `plane_b.refresh()`).
3. **Observar os logs estruturados:**
   - **Sucesso:** `whoami` responde com identidade; **não** há evento `refresh_failure`.
     → o refresh silencioso **passa** sob a CA atual. **(G3 ✓)**
   - **Falha:** surge `{"event":"refresh_failure", "reason":"invalid_grant"/..., ...}` e
     `whoami` devolve `status: reauth_required`. → a CA está a **bloquear** o refresh do
     servidor (não-gerido). **(possível N2)**

## Passo 6 — Se o refresh for bloqueado: exceção de CA de escopo mínimo  → **(G3 / R3)**

1. O admin Entra aplica a **exceção de CA de escopo mínimo**: *named location* = IP da VPS,
   **ou** a identidade do servidor (preferível). **Não** alargar para além do necessário.
2. Repetir o **Passo 5**. Registar o **antes/depois** (logs).

> **Critério (G3):** existe **pelo menos um caminho viável** em que o refresh passa, e é
> reproduzível pelos logs. Se a única exceção que o faz passar for **inaceitável** para o
> cliente por política de segurança → **R3** (decisão: aceitar reautenticação periódica ou
> não avançar). Se **nenhuma** exceção razoável o resolver (token protection/binding) → **N2**.

## Passo 7 — Testar a reautenticação graciosa

1. **Revogar** a sessão do utilizador no Entra (ou aguardar a expiração do refresh token).
2. Executar `whoami`.

> **Critério:** o Claude recebe um **pedido de re-login** (`status: reauth_required`) em vez
> de uma falha silenciosa ou um erro cru. (É o comportamento coberto pelo teste
> `test_invalid_grant_reauth_graciosa`, agora confirmado no real.)

## Passo 8 — Veredito go/no-go

Preencher mapeando observações aos critérios do [plano §1.2/1.3](plano-implementacao.md):

| Critério | Observação (preencher) | Veredito |
|---|---|---|
| **G1** — connector liga (Desktop) | Ligado com sucesso no Claude Team Desktop. DCR + login no Entra concluídos. | ✅ ok |
| **G1** — connector liga (Mobile) | Ligado com sucesso no Claude Team Mobile. | ✅ ok |
| **G2** — `whoami` end-to-end | Devolveu `marcio.martins@mobiweb.pt` / Márcio Martins via Graph `User.Read`. | ✅ ok |
| **G3** — refresh passa (direto **ou** com exceção de CA aceitável) | Testado em 2026-06-02: o refresh silencioso do servidor **manteve acesso** sob a CA atual; **sem** necessidade de exceção de CA (named location). | ✅ ok |
| Reauth graciosa (Passo 7) | Coberta por teste automático (`test_invalid_grant_reauth_graciosa`); validação manual no real ainda por fazer. | ⬜ pendente |

Anexar os excertos de log relevantes (sobretudo os eventos `refresh_failure`, se houver) e a
política de CA aplicada no Passo 6.

---

## Decisões em aberto que este runbook ajuda a fechar (v1.1 §11)

- Conditional Access: **exceção de CA** vs **reautenticação periódica** (resultado dos Passos 5–6).
- Export da política de CA exata do tenant (Pré-requisitos).
- (Em paralelo, fora deste runbook) DPA com a Anthropic + DPIA + papel jurídico.
