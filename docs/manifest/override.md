---
title: Elemento Override no arquivo de manifesto
description: O elemento Override permite que você especifique o valor de uma configuração dependendo de uma condição especificada.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4cedc68a766b599e143b9e71f272d4d2aa600be3
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485122"
---
# <a name="override-element"></a>Elemento Override

Fornece uma maneira de substituir o valor de uma configuração de manifesto, dependendo de uma condição especificada. Há três tipos de condições:

- Uma Office local diferente do padrão `LocaleToken`, chamado **LocaleTokenOverride**.
- Um padrão de suporte ao conjunto de requisitos diferente do padrão `RequirementToken` , chamado **RequirementTokenOverride**.
- A origem é diferente do padrão `Runtime`, chamado **RuntimeOverride**.

Um **elemento Override** que está dentro de um **elemento Runtime** deve ser do tipo **RuntimeOverride**.

Não há atributo `overrideType` para o elemento **Override** . A diferença é determinada pelo elemento pai e pelo tipo do elemento pai. Um **elemento Override** que está dentro de um **elemento Token** cujo `xsi:type` é `RequirementToken`, deve ser do tipo **RequirementTokenOverride**. Um **elemento Override** dentro de qualquer outro elemento pai, ou dentro de um elemento **Override** do tipo `LocaleToken`, deve ser do tipo **LocaleTokenOverride**. Para obter mais informações sobre o uso desse elemento quando ele é filho de um **elemento Token** , consulte Trabalhar com substituições [estendidas do manifesto](/office/dev/add-ins/develop/extended-overrides).

Cada tipo é descrito em seções separadas posteriormente neste artigo.

## <a name="override-element-for-localetoken"></a>Elemento Override para `LocaleToken`

Um **elemento Override** expressa uma condição e pode ser lido como um "Se ... then ..." instrução. Se o **elemento Override** for do tipo **LocaleTokenOverride**, `Locale` o atributo será a condição e `Value` o atributo será o consequente. Por exemplo, o seguinte é lido "Se a configuração de Office local for fr-fr, o nome de exibição será 'Lecteur vidéo'."

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

### <a name="syntax"></a>Sintaxe

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>Contido em

|Elemento|
|:-----|
|[CitationText](citationtext.md)|
|[Descrição](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|
|[Token](token.md)|

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Locale|cadeia de caracteres|obrigatório|Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.|
|Valor|cadeia de caracteres|obrigatório|Especifica o valor da configuração expressa para a localidade especificada.|

### <a name="examples"></a>Exemplos

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
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
```

### <a name="see-also"></a>Confira também

- [Localização para suplementos do Office](/office/dev/add-ins/develop/localization)
- [Atalhos de teclado para o SharePoint](/office/dev/add-ins/design/keyboard-shortcuts)

## <a name="override-element-for-requirementtoken"></a>Elemento Override para `RequirementToken`

Um **elemento Override** expressa uma condição e pode ser lido como um "Se ... then ..." instrução. Se o **elemento Override** for do tipo **RequirementsTokenOverride**, o elemento **filho Requirements** expressará a condição e `Value` o atributo será o conseqüente. Por exemplo, a primeira  Substituição no seguinte é ler "Se a plataforma atual dá suporte ao FeatureOne versão 1.7, use a cadeia de caracteres 'oldAddinVersion' `${token.requirements}` no lugar do token na URL do vô-vô `<ExtendedOverrides>` (em vez da cadeia de caracteres padrão 'upgrade')."

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Tipo de suplemento:** Painel de tarefas

### <a name="syntax"></a>Sintaxe

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>Contido em

|Elemento|
|:-----|
|[Token](token.md)|

### <a name="must-contain"></a>Deve conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Requisitos](requirements.md)|||x|

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Valor|cadeia de caracteres|obrigatório|Valor do token de vôvão quando a condição for atendida.|

### <a name="example"></a>Exemplo

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar quais Office e plataformas podem hospedar seu complemento](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in)
- [Atalhos de teclado para o SharePoint](/office/dev/add-ins/design/keyboard-shortcuts)

## <a name="override-element-for-runtime"></a>Elemento Override para `Runtime`

> [!IMPORTANT]
> O suporte para esse elemento foi introduzido no conjunto de requisitos de Caixa de [Correio 1.10](/office/dev/add-ins/reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10) com o [recurso de ativação baseada em evento](/office/dev/add-ins/outlook/autolaunch). Confira, [clientes e plataformas](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

Um **elemento Override** expressa uma condição e pode ser lido como um "Se ... then ..." instrução. Se o **elemento Override** for do tipo **RuntimeOverride**, `type` o atributo será a condição e `resid` o atributo será o conseqüente. Por exemplo, o seguinte é ler "Se o tipo for 'javascript', `resid` será 'JSRuntime.Url'." Outlook Desktop requer esse elemento para [manipuladores de ponto de extensão LaunchEvent](/office/dev/add-ins/reference/manifest/extensionpoint#launchevent).

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Tipo de suplemento:** Email

### <a name="syntax"></a>Sintaxe

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>Contido em

- [Tempo de execução](runtime.md)

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|**type**|string|Sim|Especifica o idioma para essa substituição. No momento, `"javascript"` é a única opção com suporte.|
|**resid**|string|Sim|Especifica o local da URL do arquivo JavaScript que deve substituir o local da URL do HTML padrão definido no elemento [Runtime](runtime.md) pai.`resid` O `resid` pode ter não mais de 32 caracteres e deve corresponder a `id` um atributo de um `Url` elemento no `Resources` elemento.|

### <a name="examples"></a>Exemplos

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configurar seu Outlook para ativação baseada em eventos](/office/dev/add-ins/outlook/autolaunch)
