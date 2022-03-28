---
title: Elemento Token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: f2064cdd6b9d0a4319ef8c01d056531811c7a8a2
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485178"
---
# <a name="token-element"></a>Elemento Token

Define um token de URL individual. Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](/office/dev/add-ins/develop/extended-overrides).

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Contido em

[Tokens](tokens.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Atributos

|Atributo|Descrição|
|:-----|:-----|
|DefaultValue|Valor padrão para esse token se nenhuma condição em qualquer elemento `<Override>` filho corresponde.|
|Nome|Nome do token. Esse nome é definido pelo usuário. O tipo do token é determinado pelo atributo type.|
|xsi:type|Define o tipo de Token. Esse atributo deve ser definido como um dos:  `"RequirementsToken"`, ou  `"LocaleToken"`.|

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