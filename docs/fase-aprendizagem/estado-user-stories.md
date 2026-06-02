# Fase Aprendizagem — Módulo de Aprendizagem de Email: estado das user stories

> **Entrega:** análise funcional + **scaffold funcional e testado**. Não liga ao Microsoft
> Graph em produção além do que já existe — a aprendizagem trabalha sobre metadados locais
> e sobre um novo store local. Aprendizagem **desligada por defeito / opt-in explícito**.

## Legenda

- ✅ feito · ⬜ pendente
- **Testado (automático):** coberto por testes unit/integração, sem rede, relógio controlável.
- **Validação manual (tenant real):** execução no tenant/VPS reais — responsabilidade do
  cliente. Pendente em todas as US (scaffold; sem novas chamadas Graph a validar).

## Tabela de estado

| US | Descrição curta | Implementado | Testado (auto) | Validação manual | Notas |
|----|-----------------|:---:|:---:|:---:|-------|
| US-L.1 | Registar comportamento (só metadados; só se opt-in) | ✅ | ✅ | ⬜ | `record_action_event` ligado aos `*_confirm` de email (send/reply/forward/move/archive/delete); defensivo (try/except, nunca quebra o email). Features cifradas. |
| US-L.2 | Gerar recomendações para um email (read-only) | ✅ | ✅ | ⬜ | `email_recommendations`; ordenadas por confiança, com `rationale`. Sem opt-in → `opt_out`. Sem chamadas Graph. |
| US-L.3 | Aceitar recomendação → token de confirmação | ✅ | ✅ | ⬜ | A sugestão traz `prepare_tool`/`prepare_params`; aceitar = chamar o `*_prepare` existente → `confirmation_token` → `*_confirm`. Sem segundo mecanismo. |
| US-L.4 | Ignorar/não agir sobre recomendações | ✅ | ✅ | ⬜ | Recomendação read-only não emite token nem persiste estado; ignorar é não-agir. Feedback explícito fica em backlog. |
| US-L.5 | Opt-in / opt-out | ✅ | ✅ | ⬜ | `learning_opt_in(enabled)`; default desligado; mensagem clara do efeito. |
| US-L.6 | Apagar histórico (esquecimento) | ✅ | ✅ | ⬜ | `learning_forget` apaga tudo do `subject`; devolve contagem. |

## Garantias transversais (verificadas por testes)

- **Opt-in obrigatório:** sem consentimento não há registo nem recomendações (`test_sem_opt_in_nao_regista_nem_recomenda`).
- **Só-metadados + cifra:** o BLOB `features_enc` não contém o assunto em claro (`test_features_cifradas_no_disco`); nunca se lê o corpo (`test_nao_usa_corpo_so_metadados`).
- **Isolamento por subject:** record/list/purge filtram por subject (`test_record_e_list_isolados_por_subject`, `test_purge_total_apaga_so_o_subject`).
- **Nunca auto-executar:** `email_recommendations` é read-only e não devolve `confirmation_token`; só o `*_prepare` chamado pelo utilizador o emite (e2e).
- **Degradação graciosa:** falha do store ao registar não quebra o move (`test_registo_de_comportamento_nunca_quebra_o_email`).
- **Explicabilidade:** cada recomendação traz `rationale` com domínio e contagem (`test_recomenda_mover_para_archive_com_rationale_e_prepare`).
- **Determinismo:** mesma entrada → mesma saída (`test_determinismo_mesma_entrada_mesma_saida`).

## Onde estão os testes

- Unit: `tests/unit/test_learning_features.py`, `test_learning_recommender.py`, `test_learning_store.py`.
- Integração (ponta-a-ponta): `tests/integration/test_learning_e2e.py` (reutiliza `tests/integration/fake_graph.py`).

Correr: `python -m pytest -q` · lint: `python -m ruff check src tests`.

## Configuração (defaults conservadores)

| Variável | Default | Significado |
|---|---|---|
| `LEARNING_ENABLED` | `false` | Aprendizagem desligada a nível de servidor (mesmo ligada, exige opt-in por utilizador). |
| `LEARNING_RETENTION_DAYS` | `180` | Retenção máxima dos eventos de comportamento. |
| `LEARNING_MIN_CONFIDENCE` | `0.5` | Limiar de confiança para apresentar uma recomendação. |
| `LEARNING_TOP_N` | `3` | Nº máximo de recomendações por pedido. |
