# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O Outlook suplemento subconjunto de API da API do JavaScript para Office inclui objetos, métodos, propriedades e suplemento de eventos que você pode usar no Outlook.

> [!NOTE]
> Esta documentação é para uma **visualização** [requisito definido](/javascript/office/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.

O conjunto de requisito de visualização inclui todos os recursos do [conjunto de requisito 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - adicionado um novo objeto que representa as propriedades de um item de compromisso ou mensagem em uma pasta compartilhada, calendário ou caixa de correio.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): um novo parâmetro `options` opcional, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - adicionado um novo método que anexa um arquivo do base64 codificação em uma mensagem ou um compromisso.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – Foi adicionada uma nova função que retorna os dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - adicionado um novo método que obtém um objeto que representa o sharedProperties de um compromisso ou item de mensagem.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - adicionado um novo enum de sinalizador de bit que especifica as permissões de representante.
- [Office.EventType](/javascript/api/office/office.eventtype) - modificado para dar suporte ao evento OfficeThemeChanged por meio da adição de `OfficeThemeChanged` entrada.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)