# <a name="highresolutioniconurl-element"></a><span data-ttu-id="30d2d-101">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="30d2d-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="30d2d-102">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.</span><span class="sxs-lookup"><span data-stu-id="30d2d-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="30d2d-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="30d2d-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="30d2d-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="30d2d-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="30d2d-105">Pode conter</span><span class="sxs-lookup"><span data-stu-id="30d2d-105">Can contain</span></span>

[<span data-ttu-id="30d2d-106">Override</span><span class="sxs-lookup"><span data-stu-id="30d2d-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="30d2d-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="30d2d-107">Attributes</span></span>

|<span data-ttu-id="30d2d-108">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="30d2d-108">**Attribute**</span></span>|<span data-ttu-id="30d2d-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="30d2d-109">**Type**</span></span>|<span data-ttu-id="30d2d-110">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="30d2d-110">**Required**</span></span>|<span data-ttu-id="30d2d-111">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="30d2d-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="30d2d-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="30d2d-112">DefaultValue</span></span>|<span data-ttu-id="30d2d-113">cadeia de caracteres (URL)</span><span class="sxs-lookup"><span data-stu-id="30d2d-113">string (URL)</span></span>|<span data-ttu-id="30d2d-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="30d2d-114">required</span></span>|<span data-ttu-id="30d2d-115">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="30d2d-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="30d2d-116">Comentários</span><span class="sxs-lookup"><span data-stu-id="30d2d-116">Remarks</span></span>

<span data-ttu-id="30d2d-p101">Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos**. Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="30d2d-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="30d2d-119">A imagem deve estar em um dos seguintes formatos de arquivo em uma resolução recomendada de 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="30d2d-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="30d2d-120">Para obter mais informações, consulte a seção _criar uma identidade visual consistente para seu aplicativo_ em [listagens de criar eficazes na AppSource e dentro do Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="30d2d-120">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>
