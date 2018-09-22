# <a name="outlook-add-in-api-requirement-set-15"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.5

O Outlook suplemento subconjunto de API da API do JavaScript para Office inclui objetos, métodos, propriedades e suplemento de eventos que você pode usar no Outlook.

> [!NOTE]
> Esta documentação é para um [requisito definido](/javascript/office/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisito mais recente.

## <a name="whats-new-in-15"></a>Novidades na 1.5?

O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Ele adicionou os seguintes recursos.

- Foi adicionado o suporte para [painéis de tarefas fixáveis](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).
- Foi adicionado o suporte para chamar [APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Foi adicionada a capacidade de marcar um anexo como embutido.
- Foi adicionada a capacidade de fechar um painel de tarefas ou diálogo.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adiciona um manipulador de eventos a um evento com suporte.
- Adicionado [Office.EventType](office.md#eventtype-string): Especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.
- Foi adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Obtém a URL do ponto de extremidade de REST para esta conta de email.
- Foi modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.
- Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Foi modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Um novo valor no dicionário do `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)