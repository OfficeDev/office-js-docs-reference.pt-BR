---
title: Elemento PhoneSettings no arquivo de manifesto
description: O elemento PhoneSettings especifica o local de origem e as configurações de controle que se aplicam quando o seu complemento de email é usado em um telefone.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 09b53de93c4f75694c24c8649cfcdead5b51910d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485187"
---
# <a name="phonesettings-element"></a>Elemento PhoneSettings

Especifica o local de origem e as configurações de controle aplicadas quando o seu suplemento de email é usado em um telefone.

> [!IMPORTANT]
> O `PhoneSettings` elemento está disponível apenas no Outlook na Web clássico (geralmente conectado a versões mais antigas do servidor Exchange local) e Outlook 2013 no Windows. Para dar Outlook em Android e iOS, consulte [Add-ins for Outlook Mobile](/office/dev/add-ins/outlook/outlook-mobile-addins).

**Tipo de suplemento:** Email

## <a name="syntax"></a>Sintaxe

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Contido em

[Form](form.md)

