| Classe | Campos | Descrição |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Especifica a ID de um Layout de Slide a ser usado para o novo slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Especifica a ID de um Slide Master a ser usado para o novo slide.|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Especifica se os marcadores no parágrafo estão visíveis.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Representa o formato de marcador do parágrafo.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Representa o alinhamento horizontal do parágrafo.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Retorna a coleção `SlideMaster` de objetos que estão na apresentação.|
||[categorias](/javascript/api/powerpoint/powerpoint.presentation#tags)|Retorna uma coleção de marcas anexadas à apresentação.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Exclui a forma da coleção de formas.|
||[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Retorna a formatação de preenchimento dessa forma.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Especifica a altura, em pontos, da forma.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtém a ID exclusiva da forma.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|A distância, em pontos, do lado esquerdo da forma até o lado esquerdo do slide.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Retorna a formatação de linha do objeto de forma.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Especifica o nome dessa forma.|
||[categorias](/javascript/api/powerpoint/powerpoint.shape#tags)|Retorna uma coleção de marcas na forma.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Retorna o objeto text frame de uma forma.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior do slide.|
||[tipo](/javascript/api/powerpoint/powerpoint.shape#type)|Retorna o tipo dessa forma.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Especifica a largura, em pontos, da forma.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Especifica a altura, em pontos, da forma.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Especifica a distância, em pontos, do lado esquerdo da forma até o lado esquerdo do slide.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Especifica a distância, em pontos, da borda superior da forma até a borda superior do slide.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Especifica a largura, em pontos, da forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint. GeometricShapeType, opções?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Adiciona uma forma geométrica ao slide.|
||[addLine(connectorType?: PowerPoint. ConnectorType, opções?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Adiciona uma linha ao slide.|
||[addTextBox(text: string, options?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Adiciona uma caixa de texto ao slide com o texto fornecido como o conteúdo.|
||[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Obtém o número de formas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Obtém uma forma usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Obtém uma forma usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Obtém uma forma usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Representa a cor de primeiro plano de preenchimento da forma no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Define a formatação de preenchimento de um formato com uma cor uniforme.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Especifica a porcentagem de transparência do preenchimento como um valor de 0,0 (opaco) a 1,0 (claro).|
||[tipo](/javascript/api/powerpoint/powerpoint.shapefill#type)|Retorna o tipo de preenchimento da forma.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|Representação de código de cor HTML da cor do texto (por exemplo, "#FF0000" representa vermelho).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Representa o estilo de traço da linha.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Representa o estilo de linha da forma.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Especifica a porcentagem de transparência da linha como um valor de 0,0 (opaco) a 1,0 (claro).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Especifica se a formatação de linha de um elemento de forma está visível.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Representa a espessura da linha, em pontos.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtém o layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Retorna uma coleção de formas no slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Obtém `SlideMaster` o objeto que representa o conteúdo padrão do slide.|
||[categorias](/javascript/api/powerpoint/powerpoint.slide#tags)|Retorna uma coleção de marcas no slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Adiciona um novo slide no final da coleção.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtém a ID exclusiva do layout do slide.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtém o nome do layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|Retorna uma coleção de formas no layout do slide.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Obtém o número de layouts na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Obtém um layout usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Obtém um layout usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Obtém um layout usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtém a ID exclusiva do Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtém a coleção de layouts fornecidos pelo Slide Master para slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtém o nome exclusivo do Slide Master.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|Retorna uma coleção de formas no Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Obtém o número de Slide Masters na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Obtém um Slide Master usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Obtém um Slide Master usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Obtém um Slide Master usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtém a ID exclusiva da marca.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtém o valor da marca.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Adiciona uma nova marca no final da coleção.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Exclui a marca com a `key` determinada nesta coleção.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Obtém o número de marcas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Obtém uma marca usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Obtém uma marca usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Obtém uma marca usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|As configurações de redação automáticas do quadro de texto.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Exclui todo o texto no quadro de texto.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Especifica se o quadro de texto contém texto.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Representa margem direita, em pontos, do quadro de texto.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Representa o alinhamento vertical do quadro de texto.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Determina se as linhas quebram automaticamente para caber o texto dentro da forma.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|Retorna um `ShapeFont` objeto que representa os atributos de fonte para o intervalo de texto.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Retorna um `TextRange` objeto para a subdistragem no intervalo determinado.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Representa o formato de parágrafo do intervalo de texto.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
