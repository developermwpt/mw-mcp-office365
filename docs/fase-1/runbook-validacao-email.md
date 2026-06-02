# Runbook — validação manual do Módulo Email (tenant real)

Validação ponta-a-ponta das escritas de email (US-1.1 a US-1.8) no **tenant e VPS reais**,
através do Claude. Os testes automáticos já cobrem a lógica com Graph/Entra mockados; este
runbook valida o que só o ambiente real prova: scopes/consent, refresh sob Conditional Access
(**gate G3**) e o efeito real no Exchange Online.

> Enquanto o **G3** (refresh sob CA em dispositivo gerido) não estiver validado, o uso em
> produção fica em suspenso. Este runbook é o passo que o desbloqueia para o email.

## Pré-requisitos

1. **Scopes no Entra (app registration):** adicionar as permissões delegadas do Microsoft
   Graph `Mail.Read`, `Mail.Send`, `Mail.ReadWrite` (além das de Fase 0: `User.Read`,
   `offline_access`, `openid`, `profile`).
2. **Admin consent:** conceder consentimento de administrador para os novos scopes no tenant.
3. **VPS:** atualizar a variável `GRAPH_SCOPES` no ambiente do servidor para incluir os novos
   scopes e **reiniciar** o serviço. Confirmar que o connector reautentica (novo consent) na
   primeira utilização.
4. **Conta de teste:** uma caixa de correio do tenant onde seja seguro enviar/mover/eliminar
   (idealmente enviar para o próprio).
5. **Acesso aos logs** do servidor para inspecionar os eventos `event=audit`.

## Critério transversal — auditoria

Para **cada** operação de escrita, confirmar no log do servidor uma linha JSON com:
`"event":"audit"`, `"action":"email.<op>"`, `"outcome":"success"`, um `"subject_hash"`
(**nunca** o email/UPN em claro) e, no envio, `"recipients_count"`. **Não** deve aparecer
nenhum endereço de email nem o corpo da mensagem.

## Passos por user story

### US-1.1 — Pesquisar
- Pedir ao Claude: *"procura os meus emails com assunto que contenha 'fatura'"*.
- **Aceite se:** devolve uma lista resumida (id, assunto, remetente, preview); se houver mais
  resultados, indica que há mais páginas (`has_more`).

### US-1.2 — Ler
- Pedir: *"abre o email <id> e resume-o"*.
- **Aceite se:** o conteúdo é apresentado e eventuais scripts/conteúdo oculto não produzem
  efeito. Testar com um email que contenha HTML com `<script>` ou texto escondido com uma
  "instrução ao assistente" — o assistente **não** deve obedecer a instruções vindas do corpo.

### US-1.5 — Anexos
- Pedir: *"que anexos tem o email <id>? descarrega o primeiro"*.
- **Aceite se:** lista nome/tipo/tamanho e devolve os bytes (base64) do anexo pedido.

### US-1.3 — Enviar (prepare + confirm)
- Pedir: *"envia um email para mim próprio com assunto 'Teste MCP' e corpo 'olá'"*.
- **Aceite se:** o Claude **pede confirmação** mostrando um resumo (destinatários, assunto)
  e um token; só após confirmar é que o email chega à Caixa de Entrada. Auditoria
  `email.send`. Repetir a confirmação com o mesmo token **não** envia um segundo email
  (idempotência).

### US-1.6 — Anexo grande (>3MB)
- Pedir o envio com um anexo > 3 MB.
- **Aceite se:** o resumo indica anexo grande / envio via upload session. **Limite conhecido:**
  a transferência dos bytes em chunks ainda é TODO — validar apenas o caminho até à criação da
  sessão; para envio efetivo com anexo, usar ≤3 MB (inline).

### US-1.4 — Responder / responder-a-todos / reencaminhar
- Pedir: *"responde ao email <id> a dizer 'recebido'"*; depois *"reencaminha o email <id>
  para <outro endereço>"*.
- **Aceite se:** após confirmação, a resposta/reencaminho aparece e **mantém a thread**.
  Auditoria `email.reply` e `email.forward`. Reencaminhar **sem** destinatário deve ser
  recusado.

### US-1.7 — Mover
- Pedir: *"move o email <id> para o Arquivo"*.
- **Aceite se:** após confirmação, a mensagem deixa a Caixa de Entrada e aparece em Arquivo.
  Auditoria `email.move`.

### US-1.8 — Eliminar (soft) e eliminar permanentemente (reforçado)
- Soft: *"elimina o email <id>"* → após confirmar, vai para **Itens Eliminados**. Auditoria
  `email.delete` (`permanent:false`).
- Permanente: *"elimina permanentemente o email <id>"* → o Claude pede **confirmação
  reforçada**; uma confirmação normal **deve ser recusada** sem apagar. Só com a confirmação
  reforçada a mensagem é removida em definitivo. Auditoria `email.delete` (`permanent:true`).

## Gate G3 — refresh sob Conditional Access

Durante os passos acima, **forçar a expiração do token** (esperar ou revogar a sessão) e
repetir uma escrita. **Aceite se:** o servidor renova o token silenciosamente em dispositivo
gerido e a operação conclui **sem** novo login. Se a CA bloquear o refresh, o Claude deve
mostrar `reauth_required` (reautenticação graciosa) e **nenhuma** chamada ao Graph deve
ocorrer. Registar o resultado (sucesso/bloqueio) — é o que decide o go/no-go de produção.
