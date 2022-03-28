---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: O elemento SupportsSharedFolders define se o Outlook está disponível em pastas compartilhadas e cenários de caixa de correio compartilhadas.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf0066639b6d40bb2a460955b14b79fae6f45e8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484981"
---
# <a name="supportssharedfolders-element"></a>Elemento SupportsSharedFolders

Define se o Outlook está disponível em cenários de caixa de correio compartilhada (agora em visualização) e pastas compartilhadas (ou seja, acesso de representante). O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Ele é definido como *false* por padrão.

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.8. Confira, [clientes e plataformas](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.8](/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8)

A seguir, um exemplo do **elemento SupportsSharedFolders** .

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
