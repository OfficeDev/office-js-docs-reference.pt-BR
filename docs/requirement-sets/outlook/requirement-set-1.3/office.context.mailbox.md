---
title: Office.context.mailbox - conjunto de requisitos 1.3
description: Outlook da API de Caixa de Correio do conjunto de requisitos 1.3 do modelo de objeto Mailbox.
ms.date: 03/18/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3142e760040177ae60deee42051cfed06f19a21e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484968"
---
# <a name="mailbox-requirement-set-13"></a>mailbox (conjunto de requisitos 1.3)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/office/dev/add-ins/outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Minimum<br>nível de permissão | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | Escrever<br>Read | [Diagnóstico](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | Escrever<br>Read | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Escrever<br>Read | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | Escrever<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Métodos

| Método | Minimum<br>nível de permissão | Modos | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | Restricted | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | Restricted | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | Escrever<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | Escrever<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
