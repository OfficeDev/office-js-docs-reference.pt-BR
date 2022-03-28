---
title: Elemento GetStarted no arquivo de manifesto
description: Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote.
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: abf8c286e971f9b5455d80aadc79a1116cae350d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485195"
---
# <a name="getstarted-element"></a>Elemento GetStarted

Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Se o **elemento GetStarted** for omitido, o explicativo usará os valores dos elementos [DisplayName](displayname.md) e [Description](description.md) .

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

## <a name="child-elements"></a>Elementos filho

| Elemento                       | Obrigatório | Descrição                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Sim      | O título usado para o início do texto explicativo.     |
| [Descrição](#description)   | Sim      | A descrição / conteúdo do corpo para o texto explicativo.|
| [LearnMoreUrl](#learnmoreurl) | Sim       | Uma URL para uma página que explica o suplemento em detalhes.   |

### <a name="title"></a>Título 

Obrigatório. O título usado para o início do texto explicativo. O **atributo resid** faz referência a uma ID válida no **elemento ShortStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

### <a name="description"></a>Descrição

Obrigatório. A descrição / conteúdo do corpo para o texto explicativo. O **atributo resid** faz referência a uma ID válida no **elemento LongStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

### <a name="learnmoreurl"></a>LearnMoreUrl

Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O **atributo resid** faz referência a uma ID válida no elemento **Urls** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

> [!NOTE]
> **LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint. Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível. 

## <a name="see-also"></a>Confira também

Os exemplos de código a seguir usam o **elemento GetStarted** .

* [Suplemento Web do Excel para manipular formatação de tabelas e gráficos](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit para um Suplemento do Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
