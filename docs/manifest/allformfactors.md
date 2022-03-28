---
title: Elemento AllFormFactors no arquivo de manifesto
description: Especifica as configurações de um suplemento para todos os fatores forma.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 25d9b815cf7db440abce2c8859c9ebb15ea5df1a
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485151"
---
# <a name="allformfactors-element"></a>Elemento AllFormFactors

Especifica as configurações de um suplemento para todos os fatores forma. Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas. **AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

> [!NOTE]
> Esse elemento só é suportado em Excel no Windows, no Mac e na Web. Ele não é suportado em outros aplicativos Office ou no iOS ou Android.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Sim |  Define onde um suplemento expõe a funcionalidade. |

## <a name="allformfactors-example"></a>Exemplo de AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
