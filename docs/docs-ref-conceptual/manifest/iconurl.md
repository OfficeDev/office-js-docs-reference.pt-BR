# <a name="iconurl-element"></a>Elemento IconUrl

Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos** (Outlook) ou na interface de usuário **Configurações**  >  **Gerenciar suplementos** (Outlook Web App). Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**. Se você publicar o seu suplemento na Office Store, o ícone também será usado no site da Office Store para todos os tipos de suplementos.

A imagem deve estar em um dos seguintes formatos de arquivo: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos do painel de conteúdo e de tarefa, a imagem especificada deve ser 32 x 32 pixels. Para aplicativos de email, a imagem deve ser 64 x 64 pixels. Você também deve especificar um ícone para uso com aplicativos de host do Office em execução em altas telas DPI usando o elemento [HighResolutionIconUrl](highresolutioniconurl.md) . Para obter mais informações, consulte a seção _criar uma identidade visual consistente para seu aplicativo_ em [listagens de criar eficazes na AppSource e dentro do Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
