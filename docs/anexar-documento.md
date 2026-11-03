# Anexar Documento em Tarefa

## Objetivo

Documentar o fluxo validado para anexar um documento em uma tarefa do portal via SharePoint API.

Neste contexto, "anexar na tarefa" **não** significa usar anexo nativo do item da lista `Tarefas`.

O fluxo real do portal é:

1. descobrir os dados da tarefa no site `portaldetarefas`
2. descobrir o cliente da tarefa
3. resolver a biblioteca real do cliente
4. subir o arquivo nessa biblioteca
5. atualizar os metadados do item do arquivo
6. vincular o documento à tarefa com `ItemParentID`

## Descobertas validadas

### Lista de tarefas

- lista: `Tarefas`
- `list_guid`: `00000000-0000-0000-0000-000000000001`

Exemplo real:

- tarefa: `76245`
- título no estágio: `Execução`

### Biblioteca real do cliente

Para o cliente `ClienteAlfa`, a biblioteca validada foi:

- título exibido: `Documentos Compartilhados`
- GUID real: `00000000-0000-0000-0000-000000000002`
- pasta raiz real: `/sites/Empresa/ClienteAlfa/Documentos Compartilhados1`

Conclusão:

- não confiar só em `GetByTitle('Documentos Compartilhados')`
- a biblioteca real precisa ser resolvida pelo site e confirmada pelo `RootFolder/ServerRelativeUrl`

## Campo de vínculo confirmado

O vínculo do documento com a tarefa acontece por:

- `ItemParentID = ID da tarefa`

Exemplo real validado:

```json
{
  "ItemParentID": 76245,
  "Processo": "FPA - Admissão",
  "Area": "Depto Pessoal",
  "Competencia": "Abril 2026"
}
```

## Script

Arquivo:

- `sharepoint_task_attachment.py`

Comando:

```bash
python3 sharepoint_task_attachment.py attach \
  --task-id 76245 \
  --client-site-url "https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa" \
  --file "/caminho/local/arquivo.pdf"
```

## Variáveis de ambiente

```env
SHAREPOINT_TENANT_ID=<tenant-id>
SHAREPOINT_CLIENT_ID=<client-id>
SHAREPOINT_CERT_PFX_PATH=/caminho/para/sharepoint-tarefas-bot.pfx
SHAREPOINT_CERT_PASSWORD=<cert-password>
SHAREPOINT_SITE_URL=https://empresa.sharepoint.com/sites/Empresa/portaldetarefas
SHAREPOINT_TASK_LIST_GUID=<task-list-guid>
```

## Fluxo técnico do script

### 1. Buscar tarefa no portal

O script consulta a lista `Tarefas` no site:

- `https://empresa.sharepoint.com/sites/Empresa/portaldetarefas`

Campos mínimos usados:

- `ID`
- `Title`
- `Cliente`
- `Processo`
- `Area`
- `Competencia`

### 2. Resolver a biblioteca do cliente

O script lista as bibliotecas do site do cliente e filtra:

- `BaseTemplate = 101`
- `Title = BibliotecaPublicacao`
- `RootFolder/ServerRelativeUrl` dentro do site do cliente

### 3. Fazer upload do arquivo

Endpoint usado:

```text
POST {site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_server_relative_url}')/Files/add(url='{file_name}',overwrite=true)
```

### 4. Localizar o item do arquivo

Endpoint usado:

```text
GET {site_url}/_api/web/GetFileByServerRelativeUrl('{file_server_relative_url}')/ListItemAllFields?$select=ID,Id,__metadata/type
```

### 5. Atualizar os metadados

Endpoint usado:

```text
POST {site_url}/_api/web/GetFileByServerRelativeUrl('{file_server_relative_url}')/ListItemAllFields
```

Headers usados:

```http
X-HTTP-Method: MERGE
IF-MATCH: *
Accept: application/json;odata=verbose
Content-Type: application/json;odata=verbose
```

Payload aplicado:

```json
{
  "__metadata": {
    "type": "SP.Data.Documentos_x0020_Compartilhados1Item"
  },
  "ItemParentID": 76245,
  "Processo": "FPA - Admissão",
  "Area": "Depto Pessoal",
  "Competencia": "Abril 2026"
}
```

## Resultado real validado

Teste feito com:

- tarefa: `76245`
- arquivo: `sharepoint-attach-test.txt`

Resultado:

```json
{
  "task_id": 76245,
  "task_title": "Execução",
  "client_site_url": "https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa",
  "library_id": "00000000-0000-0000-0000-000000000002",
  "library_title": "Documentos Compartilhados",
  "library_root_folder": "/sites/Empresa/ClienteAlfa/Documentos Compartilhados1",
  "uploaded_file_server_relative_url": "/sites/Empresa/ClienteAlfa/Documentos Compartilhados1/sharepoint-attach-test.txt",
  "uploaded_item_id": 383
}
```

## Cuidados

- o anexo da tarefa é um documento em biblioteca, não um `AttachmentFiles` do item `Tarefas`
- `ItemParentID` é a chave do vínculo
- a biblioteca real pode não coincidir com o nome exibido
- o script já trata espaços no caminho do SharePoint
