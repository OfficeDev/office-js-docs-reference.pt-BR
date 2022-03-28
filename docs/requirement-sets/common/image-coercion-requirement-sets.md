---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de Coerção de Imagem com Office de Excel, PowerPoint e Word.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: c6b04847f1a46e0205cf9699f37eeec803252021
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484925"
---
# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de Coerção de Imagens

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 permite a conversão para uma imagem (`Office.CoercionType.Image`) ao escrever dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) método. Os aplicativos a seguir são suportados.

- Excel 2013 e posterior em Windows
- Excel 2016 e posterior no Mac
- Excel no iPad
- OneNote Online
- PowerPoint 2013 e posterior em Windows
- PowerPoint 2016 e posterior no Mac
- PowerPoint Online
- PowerPoint no iPad
- Word 2013 e posterior no Windows
- Word 2016 e posterior no Mac
- Word Online
- Word no iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 permite a conversão para o formato SVG (`Office.CoercionType.XmlSvg`) ao escrever dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) método. Os aplicativos a seguir são suportados.

- Excel 2021 e posterior em Windows
- Excel 2021 e posterior no Mac
- PowerPoint 2021 e posterior em Windows
- PowerPoint 2021 e posterior no Mac
- PowerPoint Online
- Word 2021 e posterior Windows
- Word 2021 e posterior no Mac

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
