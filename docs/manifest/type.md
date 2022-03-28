---
title: Elemento Type no arquivo de manifesto
description: O elemento Type especifica se o complemento equivalente é um complemento COM ou um XLL.
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38063472e403e118baa80ee854f49a1093c1dd3d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484939"
---
# <a name="type-element"></a>Elemento Type

Especifica se o complemento equivalente é um complemento COM ou um XLL.

**Tipo de complemento:** Painel de tarefas, função Personalizada

## <a name="syntax"></a>Sintaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contido em

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valores de tipo de complemento

Você deve especificar um dos seguintes valores para o `Type` elemento.

- COM: Especifica que o complemento equivalente é um complemento COM.
- XLL: especifica que o complemento equivalente é um Excel XLL.

## <a name="see-also"></a>Confira também

- [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Torne o seu suplemento do Office compatível com um suplemento COM existente](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)