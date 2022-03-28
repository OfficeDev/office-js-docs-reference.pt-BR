---
title: Elemento Tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 75512fbf230da6bf60074094446f25f79f01a2f6
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484942"
---
# <a name="tokens-element"></a>Elemento Tokens

Define tokens que podem ser usados em URLs de modelo. Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](/office/dev/add-ins/develop/extended-overrides).

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>Contido em

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>Deve conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Token](token.md)|||x|

## <a name="example"></a>Exemplo

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```