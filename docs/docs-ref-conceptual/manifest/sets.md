# <a name="sets-element"></a>Elemento Sets

Especifica o subconjunto mínimo da API do JavaScript para Office que o Suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Contidos em

[Requirements](requirements.md)

## <a name="can-contain"></a>Pode conter

[Definir](set.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|opcional|Especifica o valor padrão do atributo  **MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".|

## <a name="remarks"></a>Comentários

Para obter mais informações sobre os conjuntos de requisito, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

