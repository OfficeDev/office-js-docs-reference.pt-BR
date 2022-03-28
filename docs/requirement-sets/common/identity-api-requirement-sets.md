---
title: Conjuntos de requisitos da API de Identidade
description: Informações do conjunto de requisitos da API de identidade para Office de complementos.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 49579536fa085a6e5ebe95db350da9508798a95c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484871"
---
# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos da API de Identidade

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de identidade, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  | Office 2021 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Build 16.0.14326.20454 ou posterior | Versão 2008 (build 13127.20000) ou posterior | Sem suporte | 16.40 ou posterior | Microsoft Office SharePoint Online e OneDrive\* |

\*Atualmente, o conjunto de requisitos é suportado Office na Web apenas para documentos que são abertos Microsoft Office SharePoint Online e OneDrive.

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook e conjuntos de requisitos de API de Identidade

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> Em um Outlook do Outlook usando a ativação baseada em eventos, a [interface OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth) é suportada no Office no Windows versão 2108 (build 14326.20258) ou posterior. O [Office. A interface Auth](/javascript/api/office/office.auth) é suportada na versão 2109 (build 14425.10000) ou posterior. Para obter mais detalhes de acordo com sua versão, consulte a página histórico de atualizações do [Office 2021](/officeupdates/update-history-office-2021) ou [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) e como encontrar sua versão do cliente [Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) e o canal de atualização.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
