# <a name="set-element"></a>Elemento Set

Especifica um conjunto de requisitos a partir da API do JavaScript para Office que o seu Suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contidos em

[Sets](sets.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Nome|string|obrigatório|O nome de um [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|opcional|Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.|

## <a name="remarks"></a>Comentários

Para obter mais informações sobre os conjuntos de requisito, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Para suplementos de email, há somente um `"Mailbox"` requisito conjunto disponível. Esse conjunto de requisito contém o subconjunto inteiro de API suportada em suplementos de email para Outlook, e você deve especificar o `"Mailbox"` requisito definir no manifesto do seu email suplemento (não é opcional conforme for o caso do conteúdo e tarefa suplementos do painel). Além disso, você não pode declarar suporte para métodos específicos em suplementos de email.
