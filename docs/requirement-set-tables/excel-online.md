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
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Retorna uma coleção de guias de trabalho vinculadas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Retorna uma coleção de exibições de planilha presentes na planilha.|
