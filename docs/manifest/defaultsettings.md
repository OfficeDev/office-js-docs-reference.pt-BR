---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: f86419bf6a3c18e3ae62091c53b1e8f82c706fb1
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484893"
---
# <a name="defaultsettings-element"></a>Elemento DefaultSettings

Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentários

O local de origem e outras configurações no **elemento DefaultSettings** aplicam-se somente a complementos de conteúdo e painel de tarefas. Para os complementos de email, especifique os locais padrão para arquivos de origem e outras configurações padrão no [elemento FormSettings](formsettings.md) .
