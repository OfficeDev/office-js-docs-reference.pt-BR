| Classe | Campos | Descrição |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#cellAddress)|O endereço da célula que contém a fórmula alterada.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousFormula)|Representa a fórmula anterior, antes de ser alterada.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positionType)|A posição de inserção, na pasta de trabalho atual, das novas planilhas.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeTo)|A planilha na pasta de trabalho atual que é referenciada para o `WorksheetPositionType` parâmetro.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)|Os nomes de planilhas individuais a inserir.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#altTextDescription)|A descrição de texto alt da Tabela Dinâmica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#altTextTitle)|O título de texto alt da Tabela Dinâmica.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayBlankLineAfterEachItem_display_)|Define se uma linha em branco deve ou não ser exibida após cada item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptyCellText)|O texto que é preenchido automaticamente em qualquer célula vazia na Tabela Dinâmica se `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillEmptyCells)|Especifica se células vazias na Tabela Dinâmica devem ser preenchidas com `emptyCellText` o .|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatAllItemLabels_repeatLabels_)|Define a configuração "repetir todos os rótulos de item" em todos os campos da Tabela Dinâmica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showFieldHeaders)|Especifica se a Tabela Dinâmica exibe os headers de campo (legendas de campo e drop-downs de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshOnOpen)|Especifica se a Tabela Dinâmica é atualizada quando a workbook é aberta.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getDirectDependents__)|Retorna um objeto que representa o intervalo que contém todos os dependentes diretos de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#getExtendedRange_direction__activeCell_)|Retorna um objeto range que inclui o intervalo atual e até a borda do intervalo, com base na direção fornecida.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getMergedAreasOrNullObject__)|Retorna um objeto RangeAreas que representa as áreas mescladas nesse intervalo.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_)|Retorna um objeto range que é a célula de borda da região de dados que corresponde à direção fornecida.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize_newRange_)|Resize a tabela para o novo intervalo.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_)|Insere as planilhas especificadas de uma pasta de trabalho de origem na pasta de trabalho atual.|
||[onActivated](/javascript/api/excel/excel.workbook#onActivated)|Ocorre quando a guia de trabalho é ativada.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[tipo](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtém o tipo do evento.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|Ocorre quando uma ou mais fórmulas são alteradas nesta planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)|Ocorre quando uma ou mais fórmulas são alteradas em qualquer planilha dessa coleção.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails)|Obtém uma matriz `FormulaChangedEventDetail` de objetos, que contém os detalhes sobre todas as fórmulas alteradas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetId)|Obtém a ID da planilha na qual a fórmula foi alterada.|
