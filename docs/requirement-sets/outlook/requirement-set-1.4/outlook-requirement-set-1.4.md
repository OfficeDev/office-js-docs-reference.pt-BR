---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: Recursos e APIs que foram introduzidos para Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.4.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: dacec0f568f34c3414301177c376004099bca1cd
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484863"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.4

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-14"></a>Novidades na versão 1.4

O conjunto de requisitos 1.4 inclui todos os recursos do conjunto [de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)): exibe uma caixa de diálogo em um Office aplicativo.
- Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-messageparent-member(1)): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.
- Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog?view=outlook-js-1.4&preserve-view=true): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)) é chamado.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Conjuntos de requisitos e clientes com suporte](../outlook-api-requirement-sets.md)
