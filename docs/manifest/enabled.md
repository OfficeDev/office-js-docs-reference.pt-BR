---
title: Elemento habilitado no arquivo de manifesto
description: Saiba como especificar que um Comando de Complemento está desabilitado quando o complemento é lançado.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: c8167bde639d5d25c4a73a7112dd6efd8ded2d28
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485197"
---
# <a name="enabled-element"></a>Elemento Enabled

Especifica se um controle [Button](control-button.md) ou [Menu está](control-menu.md) habilitado quando o complemento é lançado. O **elemento Enabled** é um elemento filho de [Control](control.md). Se for omitido, o padrão será `true`.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [RibbonApi 1.0](../requirement-sets/common/ribbon-api-requirement-sets.md)

Esse elemento só é válido em Excel, PowerPoint e Word; ou seja, `Name` quando o atributo do elemento [Host](host.md) é "Workbook", "Presentation" ou "Document".

O controle pai também pode ser habilitado e desabilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](/office/dev/add-ins/design/disable-add-in-commands).

## <a name="example"></a>Exemplo

```xml
<Enabled>false</Enabled>
```
