# Fluxo Completo

## Objetivo

Consolidar o fluxo completo validado por API para o portal de tarefas:

1. criar tarefa
2. anexar documento
3. finalizar a etapa `Execucao`

## Estado validado

O fluxo abaixo foi executado com sucesso em ambiente real:

- criação da tarefa
- upload do documento na biblioteca real do cliente
- vínculo por `ItemParentID`
- transição `Execucao -> RevisaoEValidacao`

## Sequência operacional

### 1. Criar tarefa

Comando:

```bash
python3 sharepoint_task_attachment.py create \
  --area-id 3 \
  --processo-id 33 \
  --cliente-id 13 \
  --competencia "Abril 2026" \
  --prazo "2026-04-22T00:00:00" \
  --descricao "teste 4" \
  --execucao-inicio "2026-04-22T15:45:00Z" \
  --execucao-prazo "2026-04-22T03:00:00.000Z" \
  --execucao-margem "2026-04-23T03:00:00.000Z" \
  --usuario-atual-id 300349
```

Resultado real:

```json
{
  "task_id": 76247,
  "title": "Execução",
  "passo_atual": "Execucao"
}
```

### 2. Anexar documento

Comando:

```bash
python3 sharepoint_task_attachment.py attach \
  --task-id 76247 \
  --client-site-url "https://empresa.sharepoint.com/sites/Empresa/ClienteBeta" \
  --file "./sharepoint-attach-test.txt"
```

Resultado real:

```json
{
  "task_id": 76247,
  "uploaded_file_server_relative_url": "/sites/Empresa/ClienteBeta/Documentos Compartilhados1/sharepoint-attach-test.txt",
  "uploaded_item_id": 57,
  "metadata_applied": {
    "ItemParentID": 76247,
    "Processo": "FPA - Férias",
    "Area": "Depto Pessoal",
    "Competencia": "Abril 2026"
  }
}
```

### 3. Finalizar tarefa

Comando:

```bash
python3 sharepoint_task_attachment.py finalize \
  --task-id 76247
```

Resultado real:

```json
{
  "task_id": 76247,
  "from_stage": "Execucao",
  "to_stage": "RevisaoEValidacao",
  "action": "Aprovar"
}
```

## Requisitos

- autenticação app-only com certificado `.pfx`
- permissões `Sites.Read.All` e `Sites.ReadWrite.All`
- acesso ao site `portaldetarefas`
- IDs corretos de área, processo e cliente

## Documentação detalhada

- `docs/criar-tarefa.md`
- `docs/anexar-documento.md`
- `docs/finalizar-tarefa.md`

## Observações

- o anexo da tarefa não usa `AttachmentFiles` da lista `Tarefas`
- o vínculo do documento com a tarefa ocorre por `ItemParentID`
- a biblioteca real do cliente deve ser resolvida com segurança
- a finalização hoje cobre apenas a transição `Execucao -> RevisaoEValidacao`
