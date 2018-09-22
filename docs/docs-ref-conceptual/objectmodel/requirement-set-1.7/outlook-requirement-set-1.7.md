# <a name="outlook-add-in-api-requirement-set-17"></a>Requisito de Outlook suplemento API definir 1.7

O subconjunto de APIs de suplemento do Outlook para as APIs de JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

## <a name="whats-new-in-17"></a>What's new in 1.7?

Set-requisito 1.7 inclui todos os recursos do [conjunto de requisito 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Ele adicionado os recursos a seguir.

- Adicionadas novas APIs sobre o padrão de recorrência em compromissos e mensagens que são solicitações de reunião.
- Modificado a propriedade item.from para também está disponível no modo de redação.
- Adicionado suporte para eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged.

### <a name="change-log"></a>Log de alterações

- Adicionados a [partir](/javascript/api/outlook_1_7/office.from): adiciona um novo objeto que fornece um método para obter o de valor.
- Adicionado o [Organizador](/javascript/api/outlook_1_7/office.organizer): adiciona um novo objeto que fornece um método para obter o valor do organizador.
- Adicionada a [Recorrência](/javascript/api/outlook_1_7/office.recurrence): adiciona um novo objeto que fornece métodos para obter e definir o padrão de recorrência de compromissos, mas apenas obter o padrão de recorrência de mensagens que são solicitações de reunião.
- Adicionado [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): adiciona um novo objeto que representa a configuração de fuso horário do padrão de recorrência.
- Adicionado [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): adiciona um novo objeto que fornece os métodos para obter e definir as datas e horas de compromissos em uma série recorrente e para obter as datas e horas de solicitações em uma série recorrente de reunião.
- Adicionado [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): adiciona um novo método que adiciona um manipulador de eventos para um evento com suporte.
- Modificado [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): modifica para fazer o de valor no modo de redação.
- Modificado [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - modifica para obter o valor do organizador no modo de redação.
- Adicionado [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): adiciona uma nova propriedade que obtém ou define um objeto que fornece os métodos para gerenciar o padrão de recorrência de um item de compromisso. Essa propriedade também pode ser usada para obter o padrão de recorrência de uma reunião item de solicitação.
- Adicionado [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): adiciona um novo método que remove um manipulador de eventos.
- Adicionado [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): adiciona uma nova propriedade que obtém a id da série uma ocorrência pertence.
- Adicionado [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): adiciona um nova enum que especifica o dia da semana ou tipo de dia.
- Adicionado [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): adiciona um nova enum que especifica o mês.
- Adicionado [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): adiciona um nova enum que especifica o fuso horário aplicado a recorrência.
- Adicionado [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): adiciona um nova enum que especifica o tipo de recorrência.
- Adicionado [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): adiciona um nova enum que especifica a semana do mês.
- Modificado [Office.EventType](/javascript/api/office/office.eventtype): modifica para oferecer suporte a eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged por meio da adição de `RecurrenceChanged`, `RecipientsChanged`, e `AppointmentTimeChanged` entradas respectivamente.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)