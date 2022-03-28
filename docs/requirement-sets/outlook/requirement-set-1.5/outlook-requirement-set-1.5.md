---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.5.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: de9c7b8f388e969ef00c6a7e9fb9e589a9a713a9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485226"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.5

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-15"></a>Novidades na versão 1.5?

O conjunto de requisitos 1.5 inclui todos os recursos do conjunto [de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Ele adicionou os seguintes recursos.

- Adicionado suporte para [painéis de tarefas fixáveis](/office/dev/add-ins/outlook/pinnable-taskpane).
- Adicionado suporte para chamar [APIs REST](/office/dev/add-ins/outlook/use-rest-api).
- Adicionada a capacidade de marcar um anexo como embutido.
- Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.

### <a name="change-log"></a>Log de alterações

- Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): adiciona um manipulador de eventos para um evento compatível.
- Adicionado [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): remove os manipuladores de eventos para um tipo de evento com suporte.
- Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.
- Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): obtém a URL do ponto de extremidade REST para esta conta de email.
- Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.
- Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true#office-office-ui-closecontainer-member(1)).
- Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.
- Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Conjuntos de requisitos e clientes com suporte](../outlook-api-requirement-sets.md)
