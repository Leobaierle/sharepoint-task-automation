# Criar Tarefa

## Objetivo

Documentar o fluxo validado para criar uma nova tarefa no portal via SharePoint API.

## Requisição real capturada no portal

Endpoint:

```text
POST https://empresa.sharepoint.com/sites/Empresa/portaldetarefas/_api/web/lists/getbytitle('Tarefas')/items
```

Status observado:

- `201 Created`

Tipo do item:

```text
SP.Data.TarefasListItem
```

Item criado no teste manual:

- `76246`

## Payload real capturado

```json
{
  "__metadata": {
    "type": "SP.Data.TarefasListItem"
  },
  "Title": "Execução",
  "Action": "Aprovar",
  "AreaId": "3",
  "ProcessoId": "29",
  "ClienteId": "173",
  "Competencia": "Abril 2026",
  "Descricao": "teste 3 - NÃO APROVAR",
  "Prazo": "2026-04-26T00:00:00",
  "Origem": "Manual",
  "PassoAtual": "Execucao",
  "UsuarioAtualId": 300349,
  "ExecucaoInicio": "2026-04-22T15:42:28.752Z",
  "ExecucaoPrazo": "2026-04-22T03:00:00.000Z",
  "ExecucaoMargem": "2026-04-23T03:00:00.000Z",
  "Vencimento": "2026-04-22T03:00:00.000Z",
  "VencimentoMargem": "2026-04-23T03:00:00.000Z",
  "CheckListOption": false,
  "CheckListItems": "[]",
  "NotificacaoCliente": ""
}
```

## Script

Arquivo:

- `sharepoint_task_attachment.py`

Comando:

```bash
python3 sharepoint_task_attachment.py create \
  --area-id 3 \
  --processo-id 29 \
  --cliente-id 173 \
  --competencia "Abril 2026" \
  --prazo "2026-04-26T00:00:00" \
  --descricao "teste 3 - NÃO APROVAR" \
  --execucao-inicio "2026-04-22T15:42:28.752Z" \
  --execucao-prazo "2026-04-22T03:00:00.000Z" \
  --execucao-margem "2026-04-23T03:00:00.000Z" \
  --usuario-atual-id 300349
```

## Variáveis de ambiente

```env
SHAREPOINT_TENANT_ID=<tenant-id>
SHAREPOINT_CLIENT_ID=<client-id>
SHAREPOINT_CERT_PFX_PATH=/caminho/para/sharepoint-tarefas-bot.pfx
SHAREPOINT_CERT_PASSWORD=<cert-password>
SHAREPOINT_SITE_URL=https://empresa.sharepoint.com/sites/Empresa/portaldetarefas
SHAREPOINT_TASK_LIST_GUID=<task-list-guid>
SHAREPOINT_EXECUCAO_ASSIGNEE=<id-ou-nome-ou-email-opcional>
```

## Campos usados na primeira versão

### Obrigatórios no comando

- `area-id`
- `processo-id`
- `cliente-id`
- `competencia`
- `prazo`
- `descricao`
- `execucao-inicio`
- `execucao-prazo`
- `execucao-margem`

### Opcionais

- `usuario-atual-id`
- `origem`
- `vencimento`
- `vencimento-margem`
- `checklist-option`
- `checklist-items`
- `notificacao-cliente`

Se `usuario-atual-id` nao for informado e `SHAREPOINT_EXECUCAO_ASSIGNEE` estiver configurado,
o script usa esse responsavel fixo para a etapa `Execucao`.

## Observação importante

Nesta primeira implementação, os campos de prazo e margem são recebidos explicitamente.

Motivo:

- o portal já os envia prontos no payload
- ainda não foi mapeada a regra exata de cálculo desses campos

Então a criação por API já é viável, mas a automação ainda depende de informar corretamente:

- `ExecucaoPrazo`
- `ExecucaoMargem`
- `Vencimento`
- `VencimentoMargem`

## Próximo refinamento recomendado

No futuro, vale mapear se essas datas:

- são calculadas no frontend
- ou derivadas do processo/área/cliente por regra fixa

Se a regra for determinística, o script pode passar a calcular isso sozinho em vez de receber pronto no comando.
