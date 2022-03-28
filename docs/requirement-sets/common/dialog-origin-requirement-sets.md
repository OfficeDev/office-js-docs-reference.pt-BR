---
title: Conjuntos de requisitos de origem da caixa de diálogo
description: Saiba mais sobre os conjuntos de requisitos de Origem da Caixa de Diálogo.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: a9328117ec06aa3357866191dcd8109a6a76deeb
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484837"
---
# <a name="dialog-origin-requirement-sets"></a>Conjuntos de requisitos de origem da caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de Origem da Caixa de Diálogo, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version para o aplicativo Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(assinatura) |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Compilar<br>15.0.5371.1000<br>ou posterior | Compilar<br>16.0.5200.1000<br>ou posterior | Compilar<br>A definir<br>ou posterior | A definir | 2,52 ou posterior | 16,52 ou posterior | Julho de 2021 | Versão 2108<br>(Build 10377.1000)<br>ou posterior |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-origin-11"></a>Origem da caixa de diálogo 1.1

A Origem da Caixa de Diálogo 1.1 é a primeira versão da API. Ele fornece suporte para mensagens entre domínios entre uma caixa de diálogo e sua página pai. Para obter detalhes sobre essas APIs, consulte o [tópico Office.ui](/javascript/api/office/office.ui) reference.

## <a name="see-also"></a>Confira também

- [Usar a API de diálogo do Office em suplementos do Office](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
