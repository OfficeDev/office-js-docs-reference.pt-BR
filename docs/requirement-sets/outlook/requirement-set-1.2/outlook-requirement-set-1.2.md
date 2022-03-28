---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: Recursos e APIs que foram introduzidos para Outlook e as APIs javaScript Office como parte da API de Caixa de Correio 1.2.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9e97bfe706e29baa821f3ac1b7180df1d920cfa0
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484907"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.2

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-12"></a>Novidades na versão 1.2?

O conjunto de requisitos 1.2 inclui todos os recursos do conjunto [de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.
- Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.
- Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.
- Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Conjuntos de requisitos e clientes com suporte](../outlook-api-requirement-sets.md)
