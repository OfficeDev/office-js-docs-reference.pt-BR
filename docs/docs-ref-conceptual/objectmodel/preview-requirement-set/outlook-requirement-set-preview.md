# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs de JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação é para uma **visualização** [requisito definido](/javascript/office/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.

O conjunto de requisito de visualização inclui todos os recursos do [conjunto de requisito 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

- [De](/javascript/api/outlook/office.from) - adicionado um novo objeto que fornece um método para obter o de valor.
- [Organizador](/javascript/api/outlook/office.organizer) - adicionado um novo objeto que fornece um método para obter o valor do organizador.
- [Recorrência](/javascript/api/outlook/office.recurrence) - adicionado um novo objeto que fornece métodos para obter e definir o padrão de recorrência de compromissos, mas apenas obter o padrão de recorrência de mensagens que são solicitações de reunião.
- [SeriesTime](/javascript/api/outlook/office.seriestime) - adicionado um novo objeto que fornece métodos para obter e definir as datas e horas de compromissos em uma série recorrente e para obter as datas e horas de solicitações em uma série recorrente de reunião.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): um novo parâmetro `options` opcional, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - adicionado um novo método que anexa um arquivo do base64 codificação em uma mensagem ou um compromisso.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - adicionado um novo método que adiciona um manipulador de eventos para um evento com suporte.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - modificado para obter o de valor no modo de redação.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – Foi adicionada uma nova função que retorna os dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) - modificado para obter o valor do organizador no modo de redação.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - adicionou uma nova propriedade que obtém ou define um objeto que fornece os métodos para gerenciar o padrão de recorrência de um item de compromisso. Essa propriedade também pode ser usada para obter o padrão de recorrência de uma reunião item de solicitação.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - adicionado um novo método que remove um manipulador de eventos.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - adicionada uma nova propriedade que obtém a id da série uma ocorrência pertence.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - adicionado um novo enum que especifica o dia da semana ou tipo de dia.
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - adicionado um novo enum que especifica o mês.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - adicionado um novo enum que especifica o fuso horário aplicado a recorrência.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - adicionado um novo enum que especifica o tipo de recorrência.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - adicionado um novo enum que especifica a semana do mês.
- [Office.EventType](/javascript/api/office/office.eventtype) - modificado para oferecer suporte a eventos RecurrenceChanged, RecipientsChanged, AppointmentTimeChanged e OfficeThemeChanged por meio da adição de `RecurrencePatternChanged`, `RecipientsChanged`, `AppointmentTimeChanged`, e `OfficeThemeChanged` entradas respectivamente.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)