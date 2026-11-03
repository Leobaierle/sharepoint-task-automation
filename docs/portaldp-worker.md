# Worker PortalDP → tarefa no portal de tarefas

Cria, no portal de tarefas do SharePoint, a tarefa das solicitações do **PortalDP** ([payroll-ticket-portal](https://github.com/Leobaierle/payroll-ticket-portal)) concluídas pelo agente DP em origens configuradas (por exemplo, só a origem **Gama**).

## Como funciona

```
PortalDP   --grava pedido-->  tabela sharepoint_task_requests (Postgres do portal)
        ^                                        |
        |  API interna (token)                   |
        +---- worker (timer, 1 min) ----+--> SharePoint (cert .pfx)
```

1. O agente conclui a solicitação **enviando anexo** → PortalDP oculta o anexo do cliente e grava o pedido.
2. `scripts/portaldp_worker.py` lista os pedidos pendentes (`GET /interno/sharepoint/pendentes`), reivindica (`POST .../iniciar`), baixa os arquivos e devolve o resultado (`POST .../resultado`).
3. Para cada pedido: cria a tarefa (Execução = `Clientes.ResponsavelFolha`), anexa os documentos na biblioteca do cliente (`ItemParentID`) e avança para **Revisão e Validação** (responsável = `revisao_assignee` da origem).
4. Processo **9** (Rotina Mensal Folha): se já existe a tarefa da competência → em Execução, reaproveita; além da Execução, **não cria nada e avisa** o agente no chamado.

Datas (regra das tarefas manuais): prazo P = conclusão + `prazo_dias`; Execução P−4, margem P−3, vencimento P−2, margem do vencimento P. Competência = mês atual. Nomes no SharePoint: `<arquivo> - PortalDP <chamado>.<ext>` (evita sobrescrever).

## Configuração

`.env.portaldp` (não versionado, no diretório do repo, `chmod 600`):

```
PORTALDP_INTERNAL_TOKEN=<mesmo valor de /etc/portaldp/portaldp.secrets.env>
PORTALDP_URL=http://localhost:8001
```

O certificado e as demais variáveis vêm do `.env.sharepoint`, como nos outros runners.

## Operar

```bash
scripts/run_portaldp_worker.sh --dry-run   # mostra o que faria; não reivindica nem escreve no SharePoint
scripts/run_portaldp_worker.sh             # processa a fila
python3 -m unittest scripts.test_portaldp_worker
```

Timer: `scripts/systemd/portaldp-worker.{service,timer}` (copiar para `/etc/systemd/system/`, `systemctl enable --now portaldp-worker.timer`).

Erros aparecem no chamado (`/dp`) com o botão **Reenviar**. Pedido `enviando` parado há mais de 15 min volta sozinho para a fila; a retomada não recria a tarefa (o `task_id` é gravado logo após criar).
