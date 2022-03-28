---
title: Abrir conjuntos de requisitos de janela do navegador
description: Especifica quais Office e builds suportam a API openBrowserWindow.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 6e2e9cfd6eead7db1108a0fe434958e0d4a6513a
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485119"
---
# <a name="open-browser-window-api-requirement-sets"></a>Conjuntos de requisitos da API da Janela do Navegador Aberto

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

O conjunto de API OpenBrowserWindow permite que os complementos abram um navegador para realizar tarefas que nem sempre podem ser feitas no controle de webview em áreas externas dentro do próprio add-in; por exemplo, baixar um arquivo PDF quando o controle webview é fornecido por Microsoft Edge.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API OpenBrowserWindow, os aplicativos host Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

|  Conjunto de requisitos  | Office 2021 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(Conectado à assinatura do Microsoft 365) |  Office no iPad<br>(Conectado à assinatura do Microsoft 365)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web  |  Servidor do Office Online  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | Build 16.0.14326.20454 ou posterior | Versão 1810 (build 16.0.11001.20074) ou posterior | 16.0.0.0 ou posterior | 16.0.0.0 ou posterior | N/D | N/D|

> [!NOTE]
> O conjunto de requisitos OpenBrowserWindowApi só está disponível da seguinte maneira:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e com build de versões de canal de atualização para Microsoft 365 Apps](/officeupdates/update-history-microsoft365-apps-by-date)
- [Qual versão do Office estou usando?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número de com build de um Office cliente](/officeupdates/update-history-microsoft365-apps-by-date)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

O OpenBrowserWindowApi 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o [tópico de referência Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
