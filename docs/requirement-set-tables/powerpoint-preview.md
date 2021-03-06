| Classe | Campos | Descrição |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Especifica a ID de um Layout de Slide a ser usado para o novo slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Especifica a ID de um Slide Master a ser usado para o novo slide.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Retorna a coleção `SlideMaster` de objetos que estão na apresentação.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Retorna uma coleção de marcas anexadas à apresentação.|
|[Formato](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete--)|Exclui a forma da coleção de formas.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtém a ID exclusiva da forma.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Retorna uma coleção de marcas na forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Obtém o número de formas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Obtém uma forma usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Obtém uma forma usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Obtém uma forma usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtém o layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Retorna uma coleção de formas no slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Obtém `SlideMaster` o objeto que representa o conteúdo padrão do slide.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Retorna uma coleção de marcas no slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Adiciona um novo slide no final da coleção.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtém a ID exclusiva do layout do slide.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtém o nome do layout do slide.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Obtém o número de layouts na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Obtém um layout usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Obtém um layout usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Obtém um layout usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtém a ID exclusiva do Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtém a coleção de layouts fornecidos pelo Slide Master para slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtém o nome exclusivo do Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Obtém o número de Slide Masters na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Obtém um Slide Master usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Obtém um Slide Master usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Obtém um Slide Master usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtém a ID exclusiva da marca.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtém o valor da marca.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|Adiciona uma nova marca no final da coleção.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|Exclui a marca com a `key` determinada nesta coleção.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|Obtém o número de marcas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|Obtém uma marca usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|Obtém uma marca usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|Obtém uma marca usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtém os itens filhos carregados nesta coleção.|
