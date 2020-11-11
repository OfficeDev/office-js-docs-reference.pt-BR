| Classe | Campos | Descrição |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , manipulador: any, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Adiciona um manipulador de eventos a um evento com suporte.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Obtém o organizador da reunião especificada.|
||[recorrência](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Obtém ou define o padrão de recorrência de um compromisso.|
||[removeHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , opções?: Office. AsyncContextOptions, retorno de chamada?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Remove um manipulador de eventos para um tipo de evento com suporte.|
||[seriesid](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|Obtém a ID da série à qual uma instância pertence.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , manipulador: any, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Adiciona um manipulador de eventos a um evento com suporte.|
||[recorrência](/javascript/api/outlook/outlook.appointmentread#recurrence)|Obtém o padrão de recorrência de um compromisso.|
||[removeHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , opções?: Office. AsyncContextOptions, retorno de chamada?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|Remove um manipulador de eventos para um tipo de evento com suporte.|
||[seriesid](/javascript/api/outlook/outlook.appointmentread#seriesid)|Obtém a ID da série à qual uma instância pertence.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[início](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[De](/javascript/api/outlook/outlook.from)|[getasync (opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|Obtém o valor de de uma mensagem.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , manipulador: any, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Adiciona um manipulador de eventos a um evento com suporte.|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|Obtém o endereço de email do remetente de uma mensagem.|
||[removeHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , opções?: Office. AsyncContextOptions, retorno de chamada?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Remove um manipulador de eventos para um tipo de evento com suporte.|
||[seriesid](/javascript/api/outlook/outlook.messagecompose#seriesid)|Obtém a ID da série à qual uma instância pertence.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , manipulador: any, Options?: Office. AsyncContextOptions, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Adiciona um manipulador de eventos a um evento com suporte.|
||[recorrência](/javascript/api/outlook/outlook.messageread#recurrence)|Obtém o padrão de recorrência de um compromisso.|
||[removeHandlerAsync (eventType: cadeia de caracteres do Office. EventType \| , opções?: Office. AsyncContextOptions, retorno de chamada?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|Remove um manipulador de eventos para um tipo de evento com suporte.|
||[seriesid](/javascript/api/outlook/outlook.messageread#seriesid)|Obtém a ID da série à qual uma instância pertence.|
|[Organizador](/javascript/api/outlook/outlook.organizer)|[getasync (opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|Obtém o valor do organizador de um compromisso como {@link Office. EmailAddressDetails | Objeto EmailAddressDetails}|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Obtém se os destinatários no campo **Cco** foram alterados.|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Obtém se os destinatários no campo **CC** foram alterados.|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|Obtém se os participantes opcionais foram alterados.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|Obtém se os participantes necessários foram alterados.|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Obtém se os recursos foram alterados.|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|Obtém se os destinatários no campo **para** foram alterados.|
|[Recurrence](/javascript/api/outlook/outlook.recurrence)|[getasync (opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|Retorna o objeto Recurrence atual de uma série de compromissos.|
||[RecurrenceType](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|Obtém ou define as propriedades da série de compromissos recorrentes.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|Obtém ou define as propriedades da série de compromissos recorrentes.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|Obtém ou define o tipo da série de compromissos recorrentes.|
||[seriestime](/javascript/api/outlook/outlook.recurrence#seriestime)|O {@link Office. Seriestime | Seriestime} objeto permite que você gerencie as datas de início e de término da série de compromissos recorrentes e|
||[setasync (recurrencePattern: recorrência, opções?: Office. AsyncContextOptions, retorno de chamada?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|Define o padrão de recorrência de uma série de compromissos.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[recorrência](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|Representa o dia do mês.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|Representa o dia da semana ou o tipo de dia, por exemplo, dia do fim de semana versus dia da semana.|
||[durante](/javascript/api/outlook/outlook.recurrenceproperties#days)|Representa o conjunto de dias da recorrência.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|Representa seu primeiro dia da semana, caso contrário, o padrão é o valor nas configurações do usuário atual.|
||[intervalo](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Representa o período entre instâncias da mesma série recorrente.|
||[Mês](/javascript/api/outlook/outlook.recurrenceproperties#month)|Representa o mês.|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|Representa o número da semana no mês selecionado por exemplo, ' First ' para a primeira semana do mês.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Representa o nome do fuso horário da recorrência.|
||[partida](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Valor inteiro que representa a diferença em minutos entre o fuso horário local e o UTC na data em que a série de reunião começou.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration ()](/javascript/api/outlook/outlook.seriestime#getduration--)|Obtém a duração em minutos de uma instância usual em uma série de compromissos recorrentes.|
||[getenddate ()](/javascript/api/outlook/outlook.seriestime#getenddate--)|Obtém a data de término de um padrão de recorrência no seguinte|
||[getendtime ()](/javascript/api/outlook/outlook.seriestime#getendtime--)|Obtém a hora de término de uma instância de compromisso ou solicitação de reunião usual de um padrão de recorrência em qualquer fuso horário que o usuário ou|
||[getstartdate ()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|Obtém a data de início de um padrão de recorrência no seguinte|
||[getstarttime ()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|Obtém a hora de início de uma instância de compromisso comum de um padrão de recorrência em qualquer zona de tempo que o usuário/suplemento define o|
||[setduration (minutos: número)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|Define a duração de todos os compromissos em um padrão de recorrência.|
||[setenddate (Date: String)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|Define a data de término de uma série de compromissos recorrentes.|
||[setenddate (ano: número, mês: número, dia: número)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|Define a data de término de uma série de compromissos recorrentes.|
||[setstartdate (Date: String)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|Define a data de início de uma série de compromissos recorrentes.|
||[setstartdate (ano: número, mês: número, dia: número)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|Define a data de início de uma série de compromissos recorrentes.|
||[setstarttime (horas: número, minutos: número)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|Define a hora de início de todas as instâncias de uma série de compromissos recorrentes em qualquer fuso horário em que o padrão de recorrência estiver definido|
||[setstarttime (hora: cadeia de caracteres)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|Define a hora de início de todas as instâncias de uma série de compromissos recorrentes em qualquer fuso horário em que o padrão de recorrência estiver definido|
