| Classe | Campos | Descrição |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Especifica o ângulo para o qual o texto é orientado para o título do eixo do gráfico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtém os valores de uma única dimensão da série de gráficos.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtém o tipo de conteúdo do comentário.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtém a matriz que contém as IDs e as IDs de comentários `CommentDetail` de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obter `CommentDetail` a matriz que contém as IDs de comentário e as IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Ocorre quando os comentários são adicionados.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Ocorre quando comentários ou respostas em uma coleção de comentários são alterados, incluindo quando as respostas são excluídas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Ocorre quando os comentários são excluídos na coleção de comentários.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtém a matriz que contém as IDs e as IDs de comentários `CommentDetail` de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Representa a ID do comentário.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Representa as IDs das respostas relacionadas que pertencem ao comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|O tipo de conteúdo da resposta.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Define o formato culturalmente apropriado de exibição de data e hora.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtém a cadeia de caracteres usada como separador de data.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data longa.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtém a cadeia de caracteres de formato por um valor de longo tempo.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data curta.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtém a cadeia de caracteres usada como separador de tempo.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparador](/javascript/api/excel/excel.pivotdatefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotdatefilter#condition)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|If `true` , filter exclui *itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|O limite inferior do intervalo para a condição `between` de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|O limite superior do intervalo para a condição `between` de filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Para `equals` , , e condições de `before` `after` `between` filtro, indica se as comparações devem ser feitas como dias inteiros.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Define um ou mais dos PivotFilters atuais do campo e os aplica ao campo.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Limpa todos os critérios de todos os filtros do campo.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Limpa todos os critérios existentes do filtro do campo do tipo determinado (se um estiver aplicado no momento).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtém todos os filtros atualmente aplicados no campo.|
||[isFiltered(filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Verifica se há filtros aplicados no campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|O filtro de data aplicado no momento do PivotField.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|O filtro de rótulo aplicado no momento do PivotField.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|O filtro manual aplicado no momento do PivotField.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|O filtro de valor aplicado no momento do PivotField.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparador](/javascript/api/excel/excel.pivotlabelfilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotlabelfilter#condition)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|If `true` , filter exclui *itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|O limite inferior do intervalo para a condição `between` de filtro.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|A subdistragem usada para `beginsWith` , e condições de `endsWith` `contains` filtro.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|O limite superior do intervalo para a condição `between` de filtro.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Uma lista de itens selecionados para filtrar manualmente.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica se a Tabela Dinâmica permite a aplicação de vários PivotFilters em um dado PivotField na tabela.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtém o número de Tabelas Dinâmicas na coleção.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtém a primeira Tabela Dinâmica da coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtém uma Tabela Dinâmica por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotvaluefilter#condition)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|If `true` , filter exclui *itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|O limite inferior do intervalo para a condição `between` de filtro.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Especifica se o filtro é para os itens N superior/inferior, N por cento superior/inferior ou N superior/inferior.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|O número limite "N" de itens, porcentagem ou soma a ser filtrado para uma condição de filtro superior/inferior.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|O limite superior do intervalo para a condição `between` de filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nome do "valor" escolhido no campo pelo qual filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Retorna um objeto que representa o intervalo que contém todos os precedentes diretos de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtém uma coleção com escopo de Tabelas Dinâmicas que se sobrepõem ao intervalo.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para a célula que está sendo descarada.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Representa a categoria do formato de número de cada célula.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtém o número `RangeAreas` de objetos nesta coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Retorna o `RangeAreas` objeto com base na posição na coleção.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Retorna o objeto com base na ID da `RangeAreas` planilha ou no nome da coleção.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Retorna o objeto com base no nome `RangeAreas` da planilha ou na ID da coleção.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Retorna uma matriz de endereços no estilo A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Retorna o `RangeAreasCollection` objeto.|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#ranges)|Retorna intervalos que compõem esse objeto em um `RangeCollection` objeto.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtém uma coleção de propriedades personalizadas no nível da planilha.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtém a chave da propriedade personalizada.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtém ou define o valor da propriedade personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Adiciona uma nova propriedade personalizada que mapeia para a chave fornecida.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtém o número de propriedades personalizadas nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
