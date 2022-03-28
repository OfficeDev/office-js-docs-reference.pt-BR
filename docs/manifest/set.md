---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto de requisitos da API JavaScript Office que seu Office Add-in requer para ser ativado por Office ou para substituir as configurações de manifesto base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f5e8273e07e1abf959def7eac262264bc922f45
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484985"
---
# <a name="set-element"></a>Elemento Set

O significado desse elemento depende de onde ele é usado no manifesto.

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, o elemento Requisitos do vôrent **é** um filho direto do [OfficeApp](officeapp.md)), o elemento **Set** especifica [](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) um conjunto de requisitos da API JavaScript do Office que seu add-in do Office precisa para ser ativado por Office.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Como bisneto de um elemento VersionOverrides

Especifica um conjunto [](/office/dev/add-ins/develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) de requisitos da API javaScript do Office que deve ser suportado pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para que [o VersionOverrides](versionoverrides.md) entre em vigor.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

**Associado a esses conjuntos de requisitos**:

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

## <a name="syntax"></a>Sintaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contido em

[Sets](sets.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|O nome de um [conjunto de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|cadeia de caracteres|opcional|Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se for especificado no elemento [Sets](sets.md) pai.|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para obter mais informações sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, consulte [Specify which Office versions and platforms can host your add-in](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in).

