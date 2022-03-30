---
title: Elemento Item no arquivo de manifesto
description: Especifica um item em um menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b717bcb3041a0164716b50116f387f0b581f2ea
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502627"
---
# <a name="item-element"></a>Elemento Item

Especifica um item em um menu.

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

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)     | Sim |  O texto do botão. |
|  [Supertip](supertip.md)  | Sim |  A dica detalhada do botão.    |
|  [Icon](icon.md)      | Sim |  Uma imagem para o botão.         |
|  [Action](action.md)    | Sim |  Especifica a ação a ser executada. Pode haver apenas um **filho action** de um **elemento Item** .  |
|  [Enabled](enabled.md)    | Não |  Especifica se o controle está habilitado quando o complemento é lançado.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se o botão deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas. Se usado, ele deve ser o *primeiro* elemento filho. |

### <a name="label"></a>Rótulo

Especifica o texto do botão por meio de seu único atributo, **resid**, que não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no **filho ShortStrings** do elemento [Resources](resources.md) .

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

## <a name="examples"></a>Exemplos

Por exemplos, consulte [Controle do tipo Menu](control-menu.md).