# Anexar documentos em pasta de cliente via API

## Objetivo

Este documento registra somente o fluxo tecnico para enviar documentos para a pasta/biblioteca de um cliente no SharePoint via API.

O foco aqui e:

- autenticacao
- descoberta do destino correto
- upload do arquivo
- atualizacao de metadados do arquivo
- consultas uteis para validacao

Este documento nao trata de fluxo de tarefas, aprovacoes ou anexos da interface do portal.

## Ambiente

SharePoint Online da empresa, com autenticacao via `App Registration` no Microsoft Entra ID.

## Autenticacao que funcionou

### O que nao funcionou

Autenticacao app-only com `client secret` falhou para SharePoint REST com erro:

- `Unsupported app only token`

### O que funcionou

Autenticacao app-only com certificado `.pfx`.

Configuracao usada:

```env
SHAREPOINT_TENANT_ID=...
SHAREPOINT_CLIENT_ID=...
SHAREPOINT_CLIENT_SECRET=
SHAREPOINT_CERT_PFX_PATH=C:\caminho\sharepoint-tarefas-bot\certs\sharepoint-tarefas-bot.pfx
SHAREPOINT_CERT_PASSWORD=...
SHAREPOINT_SITE_URL=https://empresa.sharepoint.com/sites/Empresa/portaldetarefas
```

Permissoes exigidas no App Registration:

- `Sites.Read.All`
- `Sites.ReadWrite.All`

As duas como permissao de aplicativo, com consentimento administrativo concedido.

O certificado publico `.cer` deve ser carregado no App Registration.

## Informacoes necessarias para enviar o arquivo

Para anexar um documento na pasta/biblioteca do cliente via API, o minimo necessario e:

- `site_url` do cliente
- nome da biblioteca
- caminho relativo da pasta dentro da biblioteca, se existir
- caminho local do arquivo

Exemplo real validado:

- `site_url`: `https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa`
- biblioteca: `Documentos Compartilhados`
- pasta: vazia ou nula
- arquivo local: `C:\caminho\boleto.pdf`

## Como montar o caminho da pasta no SharePoint

O caminho usado pela API e o `server relative url`.

Exemplo:

```text
/sites/Empresa/ClienteAlfa/Documentos Compartilhados
```

Se houver subpasta configurada:

```text
/sites/Empresa/ClienteAlfa/Documentos Compartilhados/NomeDaPasta
```

## Endpoint de upload que funcionou

O upload foi feito com este padrao:

```text
POST {site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_server_relative_url}')/Files/add(url='{file_name}',overwrite=true)
```

### Exemplo real

```text
POST https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa/_api/web/GetFolderByServerRelativeUrl('/sites/Empresa/ClienteAlfa/Documentos Compartilhados')/Files/add(url='boleto.pdf',overwrite=true)
```

### Headers relevantes

```http
Authorization: Bearer <token>
Accept: application/json;odata=verbose
Content-Type: application/octet-stream
```

### Body

O corpo da requisicao e o binario puro do arquivo.

## Resposta importante do upload

Da resposta do upload, o campo mais importante e:

- `ServerRelativeUrl`

Exemplo:

```text
/sites/Empresa/ClienteAlfa/Documentos Compartilhados/boleto.pdf
```

Esse valor deve ser guardado porque ele sera usado para consultar e atualizar o item do arquivo.

## Como descobrir o item do arquivo apos o upload

Depois do upload, o arquivo pode ser acessado por:

```text
GET {site_url}/_api/web/GetFileByServerRelativeUrl('{file_server_relative_url}')/ListItemAllFields
```

### Exemplo

```text
GET https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa/_api/web/GetFileByServerRelativeUrl('/sites/Empresa/ClienteAlfa/Documentos Compartilhados/boleto.pdf')/ListItemAllFields
```

Esse endpoint retorna o item do arquivo dentro da biblioteca.

## Como descobrir o tipo real do item do arquivo

Para atualizar os metadados do arquivo via `MERGE`, foi necessario obter o tipo real do item:

```text
GET {site_url}/_api/web/GetFileByServerRelativeUrl('{file_server_relative_url}')/ListItemAllFields?$select=__metadata/type
```

Exemplo de retorno:

```json
{
  "d": {
    "__metadata": {
      "type": "SP.Data.Documentos_x0020_CompartilhadosItem"
    }
  }
}
```

Esse valor deve ser usado no `__metadata.type` do `MERGE`.

## Como atualizar metadados do arquivo

O padrao que funcionou foi:

```text
POST {site_url}/_api/web/GetFileByServerRelativeUrl('{file_server_relative_url}')/ListItemAllFields
```

Com headers:

```http
Authorization: Bearer <token>
Accept: application/json;odata=verbose
Content-Type: application/json;odata=verbose
X-HTTP-Method: MERGE
IF-MATCH: *
```

Corpo:

```json
{
  "__metadata": {
    "type": "SP.Data.Documentos_x0020_CompartilhadosItem"
  },
  "ItemParentID": 76174
}
```

## Como descobrir o tipo de um campo antes de atualizar

Para atualizar metadados com seguranca, foi importante descobrir o tipo real do campo na biblioteca.

Padrao:

```text
GET {site_url}/_api/web/lists/getbytitle('{list_title}')/fields/getbyinternalnameortitle('{field_name}')?$select=TypeAsString
```

Exemplo:

```text
GET https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa/_api/web/lists/getbytitle('Documentos%20Compartilhados')/fields/getbyinternalnameortitle('ItemParentID')?$select=TypeAsString
```

Retorno validado:

- `ItemParentID` era `Number`

Conclusao:

- esse campo precisava receber numero no `MERGE`

## Como listar os campos da biblioteca

Esse endpoint foi util para diagnostico e validacao de schema:

```text
GET {site_url}/_api/web/lists/getbytitle('{list_title}')/fields?$select=Title,InternalName,StaticName,EntityPropertyName,TypeAsString,Hidden,ReadOnlyField
```

Isso permite confirmar:

- nome interno do campo
- nome exposto pela API
- tipo do campo
- se ele esta oculto ou somente leitura

## Ponto critico: nao confiar apenas no titulo da biblioteca

Durante a validacao, ficou claro que depender apenas de:

```text
GetByTitle('Documentos Compartilhados')
```

nao era seguro.

Foi encontrado um caso em que:

- o arquivo subido existia e era recuperado normalmente por `GetFileByServerRelativeUrl`
- mas consultas por `GetByTitle('Documentos Compartilhados')` apontavam para outra lista/biblioteca

Conclusao:

- o titulo da biblioteca pode ser ambiguo nesse ambiente
- a forma mais segura e resolver a biblioteca real pelo caminho da pasta e usar o GUID da lista

## Como descobrir a biblioteca real pelo caminho da pasta

Padrao que funcionou:

```text
GET {site_url}/_api/web/GetList('{folder_server_relative_url}')?$select=Id,Title,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder
```

Exemplo:

```text
GET https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa/_api/web/GetList('/sites/Empresa/ClienteAlfa/Documentos Compartilhados')?$select=Id,Title,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder
```

Com isso, foi possivel descobrir:

- GUID real da biblioteca
- titulo real da biblioteca
- tipo de entidade da lista
- root folder real

## Como consultar documentos usando o GUID da lista

Depois de descobrir o GUID real da biblioteca, a consulta mais segura passa a ser:

```text
GET {site_url}/_api/Web/Lists(guid'{list_guid}')/Items?... 
```

em vez de:

```text
GET {site_url}/_api/Web/Lists/GetByTitle('{list_title}')/Items?... 
```

Isso evita cair em biblioteca errada quando houver ambiguidade por titulo.

## Exemplo de consulta de documentos por metadado

Exemplo validado:

```text
GET {site_url}/_api/Web/Lists(guid'{list_guid}')/Items?$filter=ItemParentID eq 76174&$select=ID,Id,ItemParentID,EncodedAbsUrl,LinkFilename,Modified,Editor/Title&$expand=Editor
```

Esse padrao foi util para confirmar se o documento enviado estava com o metadado correto e sendo encontrado pela biblioteca certa.

## Campos retornados como mais uteis na validacao

Em consultas de documentos, estes campos foram os mais uteis:

- `ID`
- `Id`
- `ItemParentID`
- `EncodedAbsUrl`
- `LinkFilename`
- `Modified`
- `Editor/Title`
- `File_x0020_Type`

## Exemplo real validado

Arquivo enviado:

- `boleto.pdf`

Biblioteca:

- `Documentos Compartilhados`

Site:

- `https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa`

Arquivo resultante:

- `/sites/Empresa/ClienteAlfa/Documentos Compartilhados/boleto.pdf`

Metadado salvo no item:

- `ItemParentID = 76174`

Consulta final encontrou o item com:

- `EncodedAbsUrl = https://empresa.sharepoint.com/sites/Empresa/ClienteAlfa/Documentos%20Compartilhados/boleto.pdf`
- `LinkFilename = boleto.pdf`
- `Editor/Title = SharePoint App`

## Cuidados importantes

### 1. O tipo do campo precisa ser respeitado

Nao assuma que um campo e texto.

Antes de atualizar:

- consulte `TypeAsString`

### 2. Nem toda biblioteca tem o mesmo schema

Um campo existente em uma biblioteca pode nao existir em outra.

Por isso, em consultas com `$select`, o ideal e:

- selecionar apenas campos realmente necessarios
- ou descobrir antes quais campos existem

### 3. Nao confie apenas no titulo da biblioteca

Em ambiente com possiveis ambiguidades, resolva a lista real pelo caminho da pasta e use o GUID.

### 4. Guarde o `ServerRelativeUrl`

Esse valor e a chave pratica para:

- consultar o arquivo
- localizar o item
- atualizar metadados

## Sequencia recomendada para uso da API

1. autenticar com certificado
2. resolver `site_url` do cliente
3. montar `folder_server_relative_url`
4. fazer upload com `Files/add`
5. capturar `ServerRelativeUrl`
6. consultar `ListItemAllFields`
7. obter `__metadata.type`
8. descobrir tipo dos campos que serao atualizados
9. fazer `MERGE` dos metadados
10. validar com consulta da biblioteca real por GUID

## Conclusao

O fluxo de API para anexar documentos na pasta/biblioteca do cliente foi validado com sucesso.

Os pontos tecnicos mais importantes para esse ambiente foram:

- usar autenticacao por certificado
- fazer upload com `GetFolderByServerRelativeUrl(...)/Files/add(...)`
- atualizar metadados via `GetFileByServerRelativeUrl(...)/ListItemAllFields`
- respeitar o tipo real dos campos
- resolver a biblioteca real pelo caminho/GUID, nao apenas pelo titulo

## Observacoes atuais

- o script usa apenas biblioteca padrao do Python + `cryptography`
- a derivacao automatica do site do cliente usa o nome do cliente; quando houver qualquer duvida, passe `--client-site-url`
- a resolucao da biblioteca hoje filtra por titulo e valida se existe uma unica biblioteca candidata no site do cliente
- se o ambiente tiver bibliotecas duplicadas com o mesmo titulo, o correto e evoluir a resolucao para uma regra mais forte baseada em root folder conhecido ou GUID persistido
- a finalizacao implementada hoje cobre a transicao `Execucao -> RevisaoEValidacao`, baseada no payload real capturado do portal
