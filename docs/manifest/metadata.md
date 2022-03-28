---
title: Elemento Metadata no arquivo de manifesto
description: O elemento Metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: b321f0059a6cfc6f2e8855f36128a076a1266755
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485136"
---
# <a name="metadata-element"></a>Elemento Metadata

Define as configurações de metadados usados por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
