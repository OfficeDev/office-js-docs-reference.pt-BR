| Classe | Campos | Descrição |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Limpa os critérios de filtro de coluna do AutoFilter.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Faz uma solicitação para quebrar os links apontando para a lista de trabalho vinculada.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|A URL original apontando para a lista de trabalho vinculada.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Faz uma solicitação para atualizar os dados recuperados da lista de trabalho vinculada.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Quebra todos os links para as guias de trabalho vinculadas.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Faz uma solicitação para atualizar todos os links da workbook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Representa o modo de atualização dos links da agenda de trabalho.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Ativa esse modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Remove o exibição de planilha da planilha.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Cria uma cópia desse exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Cria um novo exibição de planilha com o nome determinado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Sai do exibição de planilha ativa no momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Obtém a exibição de planilha ativa da planilha no momento.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Obtém o número de exibições de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Obtém uma exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Obtém uma exibição de planilha pelo índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Exclua várias linhas de uma tabela.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Exclua um número especificado de linhas de uma tabela, começando em um determinado índice.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Retorna uma coleção de guias de trabalho vinculadas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Retorna uma coleção de exibições de planilha presentes na planilha.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Ocorre quando o nome da planilha é alterado.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Ocorre quando uma planilha é movida por um usuário dentro de uma pasta de trabalho.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Ocorre quando o nome da planilha é alterado na coleção de planilhas.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada na coleção de planilhas.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Obtém a nova posição da planilha, após a movimentação.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Obtém a posição anterior da planilha, antes da movimentação.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Obtém a ID da planilha que foi movida.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Obtém o novo nome da planilha, após a alteração do nome.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Obtém o nome anterior da planilha, antes de o nome ser alterado.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Obtém a ID da planilha com o novo nome.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Obtém o tipo do evento.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Obtém a nova configuração de visibilidade da planilha, após a alteração de visibilidade.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Obtém a configuração de visibilidade anterior da planilha, antes da alteração de visibilidade.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Obtém a ID da planilha cuja visibilidade foi alterada.|
