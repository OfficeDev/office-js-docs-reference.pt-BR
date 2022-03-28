| Classe | Campos | Descrição |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|Faz uma solicitação para quebrar os links apontando para a lista de trabalho vinculada.|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|A URL original apontando para a lista de trabalho vinculada.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|Faz uma solicitação para atualizar os dados recuperados da lista de trabalho vinculada.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|Quebra todos os links para as guias de trabalho vinculadas.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|Faz uma solicitação para atualizar todos os links da workbook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|Representa o modo de atualização dos links da agenda de trabalho.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|Ativa esse modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|Remove o exibição de planilha da planilha.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|Cria uma cópia desse exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|Obtém ou define o nome do exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|Cria um novo exibição de planilha com o nome determinado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|Cria e ativa um novo modo de exibição de planilha temporária.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|Sai do exibição de planilha ativa no momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|Obtém a exibição de planilha ativa da planilha no momento.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|Obtém o número de exibições de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|Obtém uma exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|Obtém uma exibição de planilha pelo índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|Exclua várias linhas de uma tabela.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|Exclua um número especificado de linhas de uma tabela, começando em um determinado índice.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|Retorna uma coleção de guias de trabalho vinculadas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|Retorna uma coleção de exibições de planilha presentes na planilha.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|Ocorre quando o nome da planilha é alterado.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|Ocorre quando a visibilidade da planilha é alterada.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|Ocorre quando uma planilha é movida por um usuário dentro de uma pasta de trabalho.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|Ocorre quando o nome da planilha é alterado na coleção de planilhas.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|Ocorre quando a visibilidade da planilha é alterada na coleção de planilhas.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|Obtém a nova posição da planilha, após a movimentação.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|Obtém a posição anterior da planilha, antes da movimentação.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|Obtém a ID da planilha que foi movida.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|Obtém o novo nome da planilha, após a alteração do nome.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|Obtém o nome anterior da planilha, antes de o nome ser alterado.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|Obtém a ID da planilha com o novo nome.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|Obtém o tipo do evento.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|Obtém a nova configuração de visibilidade da planilha, após a alteração de visibilidade.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|Obtém a configuração de visibilidade anterior da planilha, antes da alteração de visibilidade.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|Obtém a ID da planilha cuja visibilidade foi alterada.|
