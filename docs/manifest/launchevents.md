---
title: LaunchEvents no arquivo de manifesto
description: O elemento LaunchEvents configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 02/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 762c85158390f0b541d6fcedf41d7ce49dd7f042
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485258"
---
# <a name="launchevents-element"></a>Elemento LaunchEvents

Configura seu complemento para ser ativado com base em eventos com suporte. Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento. Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch).

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="syntax"></a>Sintaxe

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a>Contido em

[ExtensionPoint](extensionpoint.md) (**complemento de email LaunchEvent** )

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Sim |  Mapeie o evento com suporte para sua função no arquivo JavaScript para ativação do complemento. |

## <a name="see-also"></a>Confira também

- [LaunchEvent](launchevent.md)
