| Classe | Campos | Descrição |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addFileAttachmentAsync (URI: String, AttachmentName: String, Options?: Office. AsyncContextOptions & {IsInline: Boolean}, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.|
||[addItemAttachmentAsync (itemId: any, AttachmentName: String, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.|
||[body](/javascript/api/outlook/outlook.appointmentcompose#body)|Obtém um objeto que fornece métodos para manipular o corpo de um item.|
||[isInline](/javascript/api/outlook/outlook.appointmentcompose#isinline)||
||[removeAttachmentAsync (AttachmentId: cadeia de caracteres, opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|Remove um anexo de uma mensagem ou de um compromisso.|
|[AppointmentForm](/javascript/api/outlook/outlook.appointmentform)|[body](/javascript/api/outlook/outlook.appointmentform#body)|Obtém um objeto que fornece métodos para manipular o corpo de um item.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[body](/javascript/api/outlook/outlook.appointmentread#body)|Obtém um objeto que fornece métodos para manipular o corpo de um item.|
|[AttachmentDetails](/javascript/api/outlook/outlook.attachmentdetails)|[attachmentType](/javascript/api/outlook/outlook.attachmentdetails#attachmenttype)|Obtém um valor que indica o tipo de um anexo.|
||[contentType](/javascript/api/outlook/outlook.attachmentdetails#contenttype)|Obtém o tipo de conteúdo MIME do anexo.|
||[id](/javascript/api/outlook/outlook.attachmentdetails#id)|Obtém a ID do Exchange do anexo.|
||[isInline](/javascript/api/outlook/outlook.attachmentdetails#isinline)|Obtém um valor que indica se o anexo deve ser exibido no corpo do item.|
||[name](/javascript/api/outlook/outlook.attachmentdetails#name)|Obtém o nome de arquivo do anexo.|
||[size](/javascript/api/outlook/outlook.attachmentdetails#size)|Obtém o tamanho do anexo em bytes.|
|[Body](/javascript/api/outlook/outlook.body)|[getTypeAsync (opções?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult<Office. CoercionType>) => void)](/javascript/api/outlook/outlook.body#gettypeasync-options--callback--asyncresult-)|Obtém um valor que indica se o conteúdo está em formato HTML ou texto.|
||[prependAsync (data: String, Options?: Office. AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#prependasync-data--options--callback--asyncresult-)|Adiciona o conteúdo especificado ao início do corpo do item.|
||[setSelectedDataAsync (data: String, Options?: Office. AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setselecteddataasync-data--options--callback--asyncresult-)|Substitui a seleção no corpo pelo texto especificado.|
|[Location](/javascript/api/outlook/outlook.location)|[getasync (retorno de chamada: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.location#getasync-callback--asyncresult-)|Obtém o local de um compromisso.|
||[getasync (opções: Office. AsyncContextOptions, retorno de chamada: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.location#getasync-options--callback--asyncresult-)|Obtém o local de um compromisso.|
||[setasync (local: cadeia de caracteres, opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.location#setasync-location--options--callback--asyncresult-)|Define o local de um compromisso.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addFileAttachmentAsync (URI: String, AttachmentName: String, Options?: Office. AsyncContextOptions & {IsInline: Boolean}, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.|
||[addItemAttachmentAsync (itemId: any, AttachmentName: String, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.|
||[bcc](/javascript/api/outlook/outlook.messagecompose#bcc)|Obtém um objeto que fornece métodos para obter ou atualizar os destinatários na linha **Cco** (com cópia oculta) de uma mensagem.|
||[body](/javascript/api/outlook/outlook.messagecompose#body)|Obtém um objeto que fornece métodos para manipular o corpo de um item.|
||[isInline](/javascript/api/outlook/outlook.messagecompose#isinline)||
||[removeAttachmentAsync (AttachmentId: cadeia de caracteres, opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|Remove um anexo de uma mensagem ou de um compromisso.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[body](/javascript/api/outlook/outlook.messageread#body)|Obtém um objeto que fornece métodos para manipular o corpo de um item.|
|[PhoneNumber](/javascript/api/outlook/outlook.phonenumber)|[type](/javascript/api/outlook/outlook.phonenumber#type)|Obtém uma cadeia de caracteres que identifica o tipo de número de telefone: casa, trabalho, celular, não especificado.|
|[Destinatários](/javascript/api/outlook/outlook.recipients)|[addasync (Recipients: (String \| EmailUser \| EmailAddressDetails) [], Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recipients#addasync-recipients-)|Adiciona uma lista de destinatários aos destinatários existentes para um compromisso ou uma mensagem.|
||[getasync (retorno de chamada: (asyncResult: Office. AsyncResult<EmailAddressDetails [] >) => void)](/javascript/api/outlook/outlook.recipients#getasync-callback--asyncresult-)|Obtém uma lista de destinatários para um compromisso ou uma mensagem.|
||[getasync (opções: Office. AsyncContextOptions, retorno de chamada: (asyncResult: Office. AsyncResult<EmailAddressDetails [] >) => void)](/javascript/api/outlook/outlook.recipients#getasync-options--callback--asyncresult-)|Obtém uma lista de destinatários para um compromisso ou uma mensagem.|
||[setasync (Recipients: (String \| EmailUser \| EmailAddressDetails) [], callback: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|Define uma lista de destinatários para um compromisso ou uma mensagem.|
||[setasync (destinatários: (String \| EmailUser \| EmailAddressDetails) [], opções: Office. AsyncContextOptions, retorno de chamada: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|Define uma lista de destinatários para um compromisso ou uma mensagem.|
|[Assunto](/javascript/api/outlook/outlook.subject)|[getasync (retorno de chamada: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.subject#getasync-callback--asyncresult-)|Obtém o assunto de um compromisso ou uma mensagem.|
||[getasync (opções: Office. AsyncContextOptions, retorno de chamada: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.subject#getasync-options--callback--asyncresult-)|Obtém o assunto de um compromisso ou uma mensagem.|
||[setasync (Subject: String, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.subject#setasync-subject--options--callback--asyncresult-)|Define o assunto de um compromisso ou uma mensagem.|
|[Time](/javascript/api/outlook/outlook.time)|[getasync (retorno de chamada: (asyncResult: Office. AsyncResult <Date> ) => void)](/javascript/api/outlook/outlook.time#getasync-callback--asyncresult-)|Obtém a hora de início ou de término de um compromisso.|
||[getasync (opções: Office. AsyncContextOptions, retorno de chamada: (asyncResult: Office. AsyncResult <Date> ) => void)](/javascript/api/outlook/outlook.time#getasync-options--callback--asyncresult-)|Obtém a hora de início ou de término de um compromisso.|
||[setasync (dateTime: Date, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.time#setasync-datetime--options--callback--asyncresult-)|Define a hora de início ou de término de um compromisso.|