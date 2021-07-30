| Classe | Campos | Descrição |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Ocorre quando os dados dentro do controle de conteúdo são alterados.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Ocorre quando o controle de conteúdo é excluído.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Ocorre quando a seleção dentro do controle de conteúdo é alterada.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|O objeto que gerou o evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|O tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Exclui a parte XML personalizada.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Exclui um atributo com o nome dado do elemento identificado pelo xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Exclui o elemento identificado pelo xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Insere um atributo com o nome e o valor determinados ao elemento identificado pelo xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Insere o XML determinado no elemento pai identificado pelo xpath no índice de posição filho.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Consulta o conteúdo XML da parte XML personalizada.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtém a ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Obtém o URI do namespace da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Define o conteúdo XML completo da parte XML personalizada.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Atualiza o valor de um atributo com o nome dado do elemento identificado pelo xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Atualiza o XML do elemento identificado pelo xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Adiciona uma nova parte XML personalizada ao documento.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Obtém as partes XML personalizadas no documento.|
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Ocorre quando um controle de conteúdo é adicionado.|
||[configurações](/javascript/api/word/word.document#settings)|Obtém as configurações do complemento no documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Obtém as partes XML personalizadas no documento.|
||[configurações](/javascript/api/word/word.documentcreated#settings)|Obtém as configurações do complemento no documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Obtém o formato da imagem em linha.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|Obtém a fonte do marcador, número ou imagem no nível especificado na lista.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Obtém a representação de cadeia de caracteres codificada base64 da imagem no nível especificado na lista.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Redefine a fonte do marcador, número ou imagem no nível especificado na lista.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Define a imagem no nível especificado na lista.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Obtém os nomes de todos os indicadores ou sobrepostos ao intervalo.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|Insere um indicador no intervalo.|
|[Configuração](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Exclui a configuração.|
||[key](/javascript/api/word/word.setting#key)|Obtém a chave da configuração.|
||[value](/javascript/api/word/word.setting#value)|Obtém ou define o valor da configuração.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Cria uma nova configuração ou define uma configuração existente.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Exclui todas as configurações deste add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Obtém a contagem de configurações.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Mescla as células delimitadas inclusive por uma primeira e última célula.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Divide a célula no número especificado de linhas e colunas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Insere um controle de conteúdo na linha.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Mescla a linha em uma célula.|
