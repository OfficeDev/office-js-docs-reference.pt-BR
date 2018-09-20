# <a name="method-element"></a>Elemento Method

Especifica um método individual a partir da API do JavaScript para Office que o Suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contidos em

[Métodos](methods.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Nome|string|obrigatório|Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o método **getSelectedDataAsync**, você deve especificar `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Comentários

Os elementos de **métodos** e o **método** não são suportados pelos suplementos de email. Para obter mais informações sobre os conjuntos de requisito, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Porque não há um meio para especificar o requisito de versão mínima para métodos individuais, para certificar-se de que um método está disponível em tempo de execução, você também deve usar uma instrução **if** ao chamar o método no script de seu suplemento. Para obter mais informações sobre como fazer isso, consulte [Noções básicas sobre a API do JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

