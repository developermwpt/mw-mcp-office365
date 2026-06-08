# Changelog

Todas as alterações relevantes a este projeto são registadas aqui.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

## [Não publicado]

### Corrigido

- **O assunto de emails e eventos é agora exposto às tools MCP como `subject`**
  (antes `subject_line`). As tools `email_send_prepare`, `calendar_create_prepare`
  e `calendar_update_prepare` ignoravam silenciosamente um campo `subject`
  enviado pelo cliente, porque o schema só conhecia `subject_line` — o email/evento
  saía sem assunto. Não era uma limitação do MCP: o `subject` interno (o principal
  autenticado do Plano A) vem do token de acesso e nunca é parâmetro de tool, pelo
  que não havia colisão na fronteira. Adicionado teste de regressão ao schema das
  tools (`tests/integration/test_tool_schema_subject.py`).

### Alterado — ⚠️ mudança de contrato

- O parâmetro de assunto das tools de escrita passou de `subject_line` para
  `subject`. Clientes que ainda enviem `subject_line` deixam de ter o assunto
  aplicado — têm de passar a usar `subject`. As funções internas `run_*`
  mantêm o nome `subject_line` (a fronteira MCP mapeia `subject` → `subject_line`).
