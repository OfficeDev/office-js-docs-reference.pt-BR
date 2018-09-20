# <a name="requirements-element"></a>Elemento Requirements

Especifica o conjunto mínimo de requisitos da API JavaScript para Office ([conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o Suplemento do Office precisa ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contidos em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|**Elemento**|**Conteúdo**|**Email**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="remarks"></a>Comentários

Para obter mais informações sobre os conjuntos de requisito, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

