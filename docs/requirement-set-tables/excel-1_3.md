| Classe | Campos | Descrição |
|:---|:---|:---|
|[Associação](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Especifica a associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: \| String de intervalo, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Adiciona uma nova associação a um intervalo específico.|
||[addFromNamedItem (Name: String, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|
||[addFromSelection (BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Adiciona uma nova associação com base na seleção atual.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nome da Tabela Dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|A planilha que contém a Tabela Dinâmica atual.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Atualiza a Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Atualiza todas as tabelas dinâmicas da coleção.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|Representa as linhas visíveis do intervalo atual.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[fórmulas](/javascript/api/excel/excel.rangeview#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Representa a fórmula em notação no estilo L1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtém o intervalo pai associado à RangeView atual.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Representa o código de formato de número do Excel para determinada célula.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Representa os endereços de célula da RangeView.|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|O número de colunas visíveis.|
||[índice](/javascript/api/excel/excel.rangeview#index)|Retorna um valor que representa o índice da RangeView.|
||[Validação](/javascript/api/excel/excel.rangeview#rowcount)|O número de linhas visíveis.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Representa uma coleção de exibições de tabelas associadas ao intervalo.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valores de texto do intervalo especificado.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Representa o tipo de dados de cada célula.|
||[values](/javascript/api/excel/excel.rangeview#values)|Representa os valores brutos da exibição do intervalo especificado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtém uma linha RangeView por meio de seu índice.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Especifica se a primeira coluna contém formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Especifica se a última coluna contém formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Especifica se as colunas mostram a formatação em tiras nas quais as colunas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Especifica se as linhas mostram a formatação em tiras nas quais as linhas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Especifica se os botões de filtro estão visíveis na parte superior de cada cabeçalho de coluna.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Tabelas dinâmicas](/javascript/api/excel/excel.workbook#pivottables)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Tabelas dinâmicas](/javascript/api/excel/excel.worksheet#pivottables)|Coleção de Tabelas Dinâmicas que fazem parte da planilha.|
