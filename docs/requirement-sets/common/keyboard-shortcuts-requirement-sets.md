---
title: Conjuntos de requisitos de Atalhos de teclado
description: Informações de conjunto de requisitos de atalhos de teclado para Office de complementos.
ms.date: 02/15/2022
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 62160672ffc8517c5683f54248da9797db2205d0
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485177"
---
# <a name="keyboard-shortcuts-requirement-sets"></a>Conjuntos de requisitos de Atalhos de teclado

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de Atalhos de Teclado, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | N/D | Versão: 2111 (build 14701.10000) | N/D | 16.55 | Setembro de 2021 |

> [!NOTE]
> O **conjunto de requisitos KeyboardShortcuts 1.1** é suportado somente em Excel.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="keyboardshortcuts-11"></a>KeyboardShortcuts 1.1

Para obter detalhes sobre as APIs neste conjunto de requisitos, [consulte Office.actions](/javascript/api/office/office.actions).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
