| Classe | Campos | Descrição |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Obtém o número de associações da coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Obtém um objeto de associação pela ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Retorna o número de gráficos da planilha.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Obtém um gráfico usando o respectivo nome.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Retorna o número de pontos do gráfico da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Retorna o número de série da coleção.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[Retire](/javascript/api/excel/excel.nameditem#comment)|Especifica o comentário associado a esse nome.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Exclui o nome fornecido.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Retorna o objeto Range associado ao nome.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Especifica se o nome tem como escopo a pasta de trabalho ou uma planilha específica.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Retorna a planilha em que o item nomeado tem escopo.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Retorna a planilha em que o item nomeado tem escopo.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (Name: String, Reference: \| cadeia de caracteres de intervalo, comentário?: cadeia de caracteres)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Adiciona um novo nome à coleção do escopo fornecido.|
||[addFormulaLocal (Name: String, formula: String, comment?: String)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Obtém o número de itens nomeados na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Obtém um objeto NamedItem usando seu nome.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Retorna o intervalo usado do objeto de intervalo determinado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Obtém o número de objetos RangeView na coleção.|
|[Configuração](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Exclui a configuração.|
||[key](/javascript/api/excel/excel.setting#key)|A chave que representa a ID da configuração.|
||[value](/javascript/api/excel/excel.setting#value)|Representa o valor armazenado para esta configuração.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (Key: String, value: String \| número \| Boolean \| Data \| array <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Obtém o número de Configurações na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Obtém uma entrada de configuração por meio da tecla.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Obtém uma entrada de configuração por meio da tecla.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Ocorre quando as Configurações no documento são alteradas.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[configurações](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Obtém uma tabela pelo nome ou ID.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Obtém a quantidade de colunas na tabela.|
||[getItemOrNullObject (Key: String de número \| )](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Obtém um objeto de coluna por nome ou ID.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Obtém a quantidade de linhas na tabela.|
|[Workbook](/javascript/api/excel/excel.workbook)|[configurações](/javascript/api/excel/excel.workbook#settings)|Representa uma coleção de configurações associada à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas.|
||[das](/javascript/api/excel/excel.worksheet#names)|Coleção de nomes com escopo para a planilha atual.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetCount (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Obtém o número de planilhas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Obtém um objeto worksheet usando o Nome ou ID dele.|
