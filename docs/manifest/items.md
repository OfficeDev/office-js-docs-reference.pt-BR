---
title: Elemento Items no arquivo de manifesto
description: Especifica os itens em um menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 15702e2cdbeb6d8df6d714d59206b6e03143b6a4
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503390"
---
# <a name="items-element"></a>Elemento Items

Especifica os itens em um menu.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="syntax"></a>Sintaxe

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>Contido em

[Elemento Control do tipo Menu](control-menu.md)

## <a name="must-contain"></a>Deve conter

[Item](item.md)

## <a name="examples"></a>Exemplos

Por exemplos, consulte [Controle do tipo Menu](control-menu.md).