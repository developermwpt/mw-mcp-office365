# src — código técnico

Área do **código técnico** do servidor MCP. Mantida separada dos documentos funcionais,
que vivem em [`../docs`](../docs).

## Estrutura atual (PoC Fase 0)

```
mcp_o365/
├── config.py              # configuração/segredos (pydantic-settings, SecretStr)
├── logging_setup.py       # logging JSON; evento refresh_failure (sinal da CA)
├── app.py                 # composition root (build_app / main) + console script
├── server.py              # FastMCP: tool whoami + rotas /callback, /healthz, /readyz
├── auth/
│   ├── errors.py          # ReauthRequired / InvalidGrant / ConsentRequired / ...
│   ├── plane_b.py         # Plano B: Entra ID via msal (authcode+PKCE, refresh)
│   ├── plane_a.py         # Plano A: MwOAuthProvider + MwTokenVerifier (SDK)
│   ├── metadata.py        # adaptador AuthSettings (RFC 9728/8414 via SDK)  # NOTA SDK
│   └── dcr.py             # adaptador ClientRegistrationOptions (RFC 7591)  # NOTA SDK
├── identity/
│   ├── models.py          # McpPrincipal, GraphSession, LinkedAccount (multi-conta)
│   └── mapping.py         # subject (Plano A) -> sessão Graph (Plano B)
├── storage/
│   ├── crypto.py          # Cipher (ABC) + LocalAesGcmCipher (AES-256-GCM)
│   ├── token_store.py     # SQLite cifrado, isolado por subject
│   └── schema.sql         # DDL
├── graph/client.py        # wrapper /me (401/403/429 + Retry-After)
├── tools/whoami.py        # tool read-only + run_whoami (orquestração testável)
└── observability/health.py
```

`prompts/` mantém os assets de runtime (ex.:
[`assistant-playbook.md`](prompts/assistant-playbook.md)).

## Desenvolvimento

```bash
/opt/homebrew/bin/python3.12 -m venv .venv
.venv/bin/pip install -e ".[dev]"
.venv/bin/pytest                      # testes (usam pythonpath=src)
```

Configuração via `.env` (ver [`../.env.example`](../.env.example)); nunca commitar segredos.

> **Nota — install editável + caminho com espaços.** O caminho deste projeto contém um
> espaço (`MCP Office 365`). Nesta combinação (Python 3.12 do Homebrew, `site` *frozen*), o
> `.pth` gerado pelo `pip install -e .` — que não termina em newline — não é honrado pelo
> `site`, e `import mcp_o365` falha fora do `pytest`. Mitigações já aplicadas:
> - **Testes:** `pyproject.toml` define `pythonpath = ["src"]`, por isso o `pytest` funciona
>   sempre.
> - **Execução local fora do pytest:** usar `PYTHONPATH=src` (ex.:
>   `PYTHONPATH=src .venv/bin/python -m mcp_o365.app`).
> - **Produção (VPS):** usar install **não-editável** (`pip install .`), que copia o package
>   para `site-packages` e não usa `.pth` — sem este problema. Em alternativa, um checkout
>   num caminho sem espaços também resolve.

## Estado

PoC Fase 0 implementada (T0–T12). Pré-requisito de validação: ver
[análise funcional v1.1 §2.4 e §10](../docs/analise-funcional-v1.1.md) e o
[plano de implementação](../docs/poc-fase-0/plano-implementacao.md). A validação real do
connector e do refresh sob Conditional Access é manual (runbook), no tenant/VPS reais.
