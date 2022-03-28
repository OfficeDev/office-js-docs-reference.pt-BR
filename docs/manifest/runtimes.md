---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: b0ff18a44753c7c9d8a8f8886e40fb11c0e155d8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484876"
---
# <a name="runtimes-element"></a>Elemento Runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao executar no Office no Windows, um add-in `<Runtimes>` que tenha um elemento em seu manifesto não necessariamente é executado no mesmo controle de webview como faria. Para obter mais informações sobre como as versões de Windows e Office determinam qual controle webview normalmente é usado, consulte [Browsers used by Office Add-ins](/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins). Se as condições descritas lá para o uso de Microsoft Edge com WebView2 (Chromium baseado em Chromium) são atendidas, o complemento usa esse navegador se ele tem ou não um `<Runtimes>` elemento. No entanto, quando essas condições não são atendidas, um complemento com um elemento sempre usa o `<Runtimes>` Internet Explorer 11, independentemente da versão Windows ou Microsoft 365.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

 - Painel de tarefas 1.0
 - Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [SharedRuntime 1.1](../requirement-sets/common/shared-runtime-requirement-sets.md) (Somente quando usado em um complemento do painel de tarefas.)

[!include[Runtimes support](../includes/runtimes-note.md)]

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

[Host](host.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Tempo de execução](runtime.md) | Sim |  O tempo de execução do seu complemento. **Importante**: no momento, você só pode definir um `<Runtime>` elemento. |

## <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Configurar seu Outlook para ativação baseada em eventos](/office/dev/add-ins/outlook/autolaunch)
