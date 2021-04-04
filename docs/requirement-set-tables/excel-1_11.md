| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fornece informações com base nas configurações atuais de cultura do sistema.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtém a cadeia de caracteres usada como separador decimal para valores numéricos.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Especifica se os separadores do sistema do Excel estão habilitados.|
|[Comentário](/javascript/api/excel/excel.comment)|[menções](/javascript/api/excel/excel.comment#mentions)|Obtém as entidades (por exemplo, pessoas) mencionadas nos comentários.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtém o conteúdo rich comment (por exemplo, menções nos comentários).|
||[resolvido](/javascript/api/excel/excel.comment#resolved)|O status do thread de comentário.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Atualiza o conteúdo do comentário com uma cadeia de caracteres especialmente formatada e uma lista de menções.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Cadeia de caracteres de intervalo, conteúdo: cadeia de caracteres \| CommentRichContent, \| contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Cria um novo comentário com o conteúdo fornecido na célula especificada.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|O endereço de email da entidade mencionada em um comentário.|
||[id](/javascript/api/excel/excel.commentmention#id)|A ID da entidade.|
||[name](/javascript/api/excel/excel.commentmention#name)|O nome da entidade mencionada em um comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[menções](/javascript/api/excel/excel.commentreply#mentions)|As entidades (por exemplo, pessoas) mencionadas nos comentários.|
||[resolvido](/javascript/api/excel/excel.commentreply#resolved)|O status da resposta ao comentário.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|O conteúdo rich comment (por exemplo, menções nos comentários).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Atualiza o conteúdo do comentário com uma cadeia de caracteres especialmente formatada e uma lista de menções.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para um comentário.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[menções](/javascript/api/excel/excel.commentrichcontent#mentions)|Uma matriz que contém todas as entidades (por exemplo, pessoas) mencionadas no comentário.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Especifica o conteúdo rico do comentário (por exemplo, conteúdo de comentário com menções, a primeira entidade mencionada tem um atributo ID de 0 e a segunda entidade mencionada tem um atributo ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtém o nome da cultura no formato languagecode2-country/regioncode2 (por exemplo, "zh-cn" ou "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Define o formato culturalmente apropriado de exibição de números.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtém a cadeia de caracteres usada como separador decimal para valores numéricos.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Move valores de célula, formatação e fórmulas do intervalo atual para o intervalo de destino, substituindo as informações antigas nessas células.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajusta o recuo da formatação do intervalo.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas foi alterado em uma planilha específica.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|O endereço do intervalo que concluiu o cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas foi alterado em uma planilha específica.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento foi disparado.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtém a ID da planilha na qual os dados foram alterados.|
