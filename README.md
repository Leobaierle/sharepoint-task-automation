# Automação de tarefas no SharePoint (portal de tarefas)

CLI e workers em Python que **criam, anexam e finalizam tarefas** num portal de tarefas construído em listas do
SharePoint Online. Também geram as **tarefas mensais recorrentes** dos departamentos (folha, adiantamento,
DCTFWeb, FGTS, rotina fiscal), com prazos em dias úteis. Os agendamentos rodam no n8n e um painel acompanha as
execuções.

> Projeto real, em produção, publicado aqui sem dados de clientes. URLs de tenant, GUIDs de lista, IDs e nomes foram
> trocados por exemplos.

## Destaques

- **Autenticação app-only por certificado** (Entra ID). Com *client secret* o SharePoint REST recusa
  (`Unsupported app only token`), então o script monta e assina o *client assertion* JWT com o `.pfx` usando só
  `cryptography` e a biblioteca padrão. Sem MSAL nem Office365-REST.
- **Fila no Postgres:** `enqueue-monthly-routine` calcula competência, prazo e datas de cada etapa e grava em
  `sharepoint_task_requests`. `process-pending-creates` consome a fila, cria a tarefa, resolve o responsável pelo
  cadastro do cliente e registra o resultado. É idempotente e retomável.
- **Calendário de prazos:** dias úteis com feriados nacionais (Páscoa calculada), "N dias úteis antes do
  vencimento", margem e vencimento por rotina.
- **Planilha como fonte da verdade:** a rotina fiscal lê um `.xlsx` direto do XML (`zipfile` + `ElementTree`),
  incluindo a **cor de preenchimento** das células: vermelho sai, amarelo e azul entram. O resultado é sincronizado
  com a tabela de clientes da rotina.
- **Anexar no lugar certo:** descobre a biblioteca real do cliente pelo caminho da pasta (não pelo título, que se
  repete), sobe o arquivo e grava os metadados respeitando o tipo de cada campo (`ItemParentID` etc.).
- **Worker do portal de solicitações** (`scripts/portaldp_worker.py`): consome a API interna do
  [payroll-ticket-portal](https://github.com/Leobaierle/payroll-ticket-portal), cria a tarefa, anexa os arquivos e
  avança para Revisão e Validação. Roda num timer do `systemd` e tem testes (`unittest`, 15 casos).
- **Runners observáveis:** cada `scripts/run_*.sh` emite `@@ etapa: ...` no stderr. O painel mostra onde a execução
  parou quando algo falha.

## Comandos

```bash
python3 sharepoint_task_attachment.py create ...                 # cria tarefa na etapa Execução
python3 sharepoint_task_attachment.py attach ...                 # anexa arquivo na biblioteca do cliente
python3 sharepoint_task_attachment.py finalize ...               # Execução -> Revisão e Validação
python3 sharepoint_task_attachment.py enqueue-monthly-routine --routine-key folha_mensal --dry-run
python3 sharepoint_task_attachment.py describe-monthly-routine --routine-key folha_mensal
python3 sharepoint_task_attachment.py process-pending-creates --limit 50
```

Configuração em `.env.sharepoint` (não versionado): `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`,
`SHAREPOINT_CERT_PFX_PATH`, `SHAREPOINT_CERT_PASSWORD`, `SHAREPOINT_SITE_URL` e a conexão do Postgres.

## Documentação

- [docs/fluxo-completo.md](docs/fluxo-completo.md): criar → anexar → finalizar, com exemplos
- [docs/criar-tarefa.md](docs/criar-tarefa.md), [docs/anexar-documento.md](docs/anexar-documento.md),
  [docs/finalizar-tarefa.md](docs/finalizar-tarefa.md): payloads reais do portal
- [docs/api-rest-referencia.md](docs/api-rest-referencia.md): referência da SharePoint REST API (upload, metadados,
  tipos de campo, descoberta de biblioteca)
- [docs/portaldp-worker.md](docs/portaldp-worker.md): integração com o portal de solicitações
- `scripts/seed_exemplo.sql`: exemplo de carga da lista de clientes de uma rotina

## Stack

Python 3.12 (stdlib + `cryptography`), SharePoint REST API, Microsoft Entra ID, PostgreSQL (`psql`), n8n, systemd.
