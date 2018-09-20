# <a name="override-element"></a>Elemento Override

Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Contidos em

|**Elemento**|
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

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Locale|string|obrigatório|Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.|
|Valor|string|obrigatório|Especifica o valor da configuração expressa para a localidade especificada.|

## <a name="see-also"></a>Confira também

- [Localização para Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
