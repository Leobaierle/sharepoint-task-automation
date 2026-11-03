# Finalizar Tarefa

## Objetivo

Documentar o fluxo validado para finalizar a etapa `Execucao` de uma tarefa e avançar para `RevisaoEValidacao` via SharePoint API.

## Requisição real capturada no portal

Endpoint:

```text
POST https://empresa.sharepoint.com/sites/Empresa/portaldetarefas/_api/web/lists/getbytitle('Tarefas')/items(76174)
```

Headers relevantes:

```http
X-HTTP-Method: MERGE
IF-MATCH: *
Accept: application/json;odata=verbose
Content-Type: application/json;odata=verbose
```

Status observado:

- `204 No Content`

## Mudanças observadas no payload

Na finalização de `Execucao`, o portal alterou estes campos:

- `Action = "Aprovar"`
- `PassoAnterior = "Execucao"`
- `PassoAtual = "RevisaoEValidacao"`
- `Title = "Revisão e Validação"`
- `ExecucaoFim = agora`
- `RevisaoEValidacaoInicio = agora`
- `ExecucaoUserAprovId = usuario atual`
- `UsuarioAtualId = usuario atual`
- `Vencimento = RevisaoEValidacaoPrazo`
- `VencimentoMargem = RevisaoEValidacaoMargem`

Além disso, o portal preservou os demais campos relevantes do item.

## Script

Arquivo:

- `sharepoint_task_attachment.py`

Comando:

```bash
python3 sharepoint_task_attachment.py finalize \
  --task-id 76245
```

Se quiser enviar HTML específico em `Comments`:

```bash
python3 sharepoint_task_attachment.py finalize \
  --task-id 76245 \
  --comment-html "<div class='comment-container'>...</div>"
```

## Variáveis de ambiente

```env
SHAREPOINT_TENANT_ID=<tenant-id>
SHAREPOINT_CLIENT_ID=<client-id>
SHAREPOINT_CERT_PFX_PATH=/caminho/para/sharepoint-tarefas-bot.pfx
SHAREPOINT_CERT_PASSWORD=<cert-password>
SHAREPOINT_SITE_URL=https://empresa.sharepoint.com/sites/Empresa/portaldetarefas
SHAREPOINT_TASK_LIST_GUID=<task-list-guid>
SHAREPOINT_REVISAO_ASSIGNEE=<id-ou-nome-ou-email-opcional>
```

## Fluxo técnico do script

### 1. Buscar a tarefa com detalhes

O script busca o item completo da tarefa, incluindo:

- `__metadata.type`
- `UsuarioAtual`
- `Cliente`
- `Processo`
- `Area`
- datas e campos operacionais

### 2. Montar o payload de transição

O script preserva o objeto atual e altera os campos-chave da transição:

```json
{
  "Action": "Aprovar",
  "PassoAnterior": "Execucao",
  "PassoAtual": "RevisaoEValidacao",
  "Title": "Revisão e Validação",
  "ExecucaoFim": "2026-04-22T15:30:46Z",
  "RevisaoEValidacaoInicio": "2026-04-22T15:30:46Z",
  "ExecucaoUserAprovId": 300349,
  "UsuarioAtualId": 31
}
```

Se `SHAREPOINT_REVISAO_ASSIGNEE` estiver configurado, `UsuarioAtualId` passa a usar esse responsavel
na transicao para `RevisaoEValidacao`. Sem essa configuracao, o script preserva o usuario atual.

### 3. Fazer o MERGE no item da lista `Tarefas`

Endpoint:

```text
POST {task_site_url}/_api/web/lists/getbytitle('Tarefas')/items({task_id})
```

Com:

```http
X-HTTP-Method: MERGE
IF-MATCH: *
```

## Resultado real validado

Teste feito com:

- tarefa: `76245`

Resultado:

```json
{
  "task_id": 76245,
  "from_stage": "Execucao",
  "to_stage": "RevisaoEValidacao",
  "action": "Aprovar",
  "title": "Revisão e Validação",
  "execucao_fim": "2026-04-22T15:30:46Z",
  "revisao_inicio": "2026-04-22T15:30:46Z",
  "usuario_atual_id": 300349
}
```

## Pré-requisito operacional

Antes da finalização, foi confirmado que:

- o anexo enviado por API apareceu corretamente no portal
- não foi necessário salvar manualmente antes de finalizar

## Cuidados

- esta implementação cobre a transição `Execucao -> RevisaoEValidacao`
- para outros estágios, o ideal é capturar o payload real correspondente antes de automatizar
- `Comments` no portal pode carregar HTML formatado; se isso for importante para o fluxo operacional, enviar via `--comment-html`
