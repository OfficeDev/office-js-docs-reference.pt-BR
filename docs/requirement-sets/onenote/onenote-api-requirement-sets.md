---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do OneNote.
ms.date: 08/24/2020
ms.prod: onenote
ms.localizationpriority: high
ms.openlocfilehash: f2f8f765f5f8537470977d55286216ba8328efb5
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484915"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do OneNote

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos do cliente Office que oferecem suporte a esse conjunto de requisitos, e as versões de compilação ou data de disponibilidade.

|  Conjunto de requisitos  |  Office na Web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | Setembro de 2016 |  

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

A API JavaScript do OneNote 1.1 é a primeira versão da API. Para obter detalhes sobre a API, confira a [visão geral da programação da API JavaScript do OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Verificação do suporte a requisitos de tempo de execução

Durante o tempo de execução, os suplementos podem verificar se um determinado aplicativo do Office oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação:

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Verificação de suporte a requisitos com base em manifesto

Use o elemento `Requirements` no manifesto do suplemento para especificar os conjuntos de requisitos críticos ou os membros da API que seu suplemento deve usar. Se o aplicativo ou a plataforma do Office não for compatível com os conjuntos de requisitos ou membros da API especificados no elemento `Requirements`, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido em Meus Suplementos.

O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos do cliente Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](../common/office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do OneNote](/javascript/api/onenote)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
