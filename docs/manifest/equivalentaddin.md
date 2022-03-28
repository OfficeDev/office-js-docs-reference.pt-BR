---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 154e69367b00a49c896d9fb108d621ba212c4580
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484854"
---
# <a name="equivalentaddin-element"></a>Elemento EquivalentAddin

Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Tipo de complemento:** Painel de tarefas, Email, Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="syntax"></a>Sintaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contido em

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Deve conter

[Tipo](type.md)

## <a name="can-contain"></a>Pode conter

[ProgId](progid.md)
 [FileName](filename.md)

## <a name="remarks"></a>Comentários

Para especificar um complemento COM como o complemento equivalente, forneça os elementos `ProgId` e `Type` . Para especificar uma XLL como o complemento equivalente, forneça os elementos `FileName` e `Type` .

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)