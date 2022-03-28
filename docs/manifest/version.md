---
title: Elemento Version no arquivo de manifesto
description: O elemento Version especifica sua Office versão do add-in.
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484839"
---
# <a name="version-element"></a>Elemento Version

Especifica a versão de seu Complemento do Office. O número da versão pode ser 1, 2, 3 ou 4 partes (ou seja, n, n.n, n.n.n ou n.n.n.n).

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

Cada parte do número da versão pode ter no máximo 5 dígitos.
