| Classe | Campos | Descrição |
|:---|:---|:---|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[displayReplyAllFormAsync (formData: String \| ReplyFormData, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#displayreplyallformasync-formdata--options--callback--asyncresult-)|Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do|
||[displayReplyFormAsync (formData: String \| ReplyFormData, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#displayreplyformasync-formdata--options--callback--asyncresult-)|Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.|
|[Body](/javascript/api/outlook/outlook.body)|[appendOnSendAsync (data: String, Options?: Office. AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#appendonsendasync-data--options--callback--asyncresult-)|Anexa em enviar o conteúdo especificado para o final do corpo do item, após qualquer assinatura.|
|[CustomProperties](/javascript/api/outlook/outlook.customproperties)|[getAll()](/javascript/api/outlook/outlook.customproperties#getall--)|Retorna um objeto com todas as propriedades personalizadas em uma coleção de pares de nome/valor.|
|[Caixa de Correio](/javascript/api/outlook/outlook.mailbox)|[displayAppointmentFormAsync (itemId: String, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#displayappointmentformasync-itemid--options--callback--asyncresult-)|Exibe um compromisso de calendário existente.|
||[displayMessageFormAsync (itemId: String, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#displaymessageformasync-itemid--options--callback--asyncresult-)|Exibe uma mensagem existente.|
||[displayNewAppointmentFormAsync (parâmetros: AppointmentForm, opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#displaynewappointmentformasync-parameters--options--callback--asyncresult-)|Exibe um formulário para criar um compromisso no calendário.|
||[displayNewMessageFormAsync (parâmetros: any, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#displaynewmessageformasync-parameters--options--callback--asyncresult-)|Exibe um formulário para criar uma nova mensagem.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[displayReplyAllFormAsync (formData: String \| ReplyFormData, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#displayreplyallformasync-formdata--options--callback--asyncresult-)|Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do|
||[displayReplyFormAsync (formData: String \| ReplyFormData, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#displayreplyformasync-formdata--options--callback--asyncresult-)|Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.|