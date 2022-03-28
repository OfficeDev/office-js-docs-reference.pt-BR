---
title: Elemento EquivalentAddins no arquivo de manifesto
description: Especifica a compatibilidade com compatibilidade com um complemento COM equivalente, XLL ou ambos.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 406fb43b74afa31b6a325f7e3d3648d89093aadc
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484999"
---
# <a name="equivalentaddins-element"></a>Elemento EquivalentAddins

Especifica a compatibilidade com compatibilidade com um complemento COM equivalente, XLL ou ambos.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Tipo de complemento:** Painel de tarefas, Email, Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>Contido em

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>Deve conter

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)