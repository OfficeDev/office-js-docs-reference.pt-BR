---
title: Elemento Namespace no arquivo de manifesto
description: O elemento Namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf05f7c950e2be03baf056284bdcb0bd6a5b5d0e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485257"
---
# <a name="namespace-element"></a>Elemento Namespace

Define o namespace usado por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Não  | Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md). Não pode ter mais de 32 caracteres. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<Namespace resid="namespace" />
```
