---
title: Elemento Sets no arquivo de manifesto
description: O elemento Sets especifica o conjunto mínimo de api JavaScript Office que seu Office Add-in requer para ser ativado pelo Office ou para substituir as configurações de manifesto base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 841b867ef4f40c63e114cf9b1dd5ce7ae245690c
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64501766"
---
# <a name="sets-element"></a>Elemento Sets

O significado desse elemento depende de onde ele é usado no manifesto.

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, o elemento **Requirements** pai é um filho direto do [OfficeApp](officeapp.md)), o elemento **Sets** especifica o subconjunto mínimo dos requisitos de API JavaScript [do Office](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) (conjuntos de requisitos) que seu Office Add-in precisa para ser ativado pelo Office.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>Como um neto de um elemento VersionOverrides

Especifica o conjunto mínimo de requisitos de API JavaScript [do Office](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) (conjuntos de requisitos) que devem ser suportados pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para que [o VersionOverrides](versionoverrides.md) entre em vigor.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- O mesmo que o elemento [Pai Requirements](requirements.md) .

**Associado a esses conjuntos de requisitos**:

- O mesmo que o elemento [Pai Requirements](requirements.md) .

## <a name="syntax"></a>Sintaxe

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Contido em

[Requisitos](requirements.md)

## <a name="can-contain"></a>Pode conter

[Set](set.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|cadeia de caracteres|opcional|Especifica o valor padrão **do atributo MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para obter mais informações sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, consulte [Specify which Office versions and platforms can host your add-in](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in).

