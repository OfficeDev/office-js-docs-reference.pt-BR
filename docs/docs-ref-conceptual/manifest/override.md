# <a name="override-element"></a><span data-ttu-id="0b79e-101">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="0b79e-101">Override element</span></span>

<span data-ttu-id="0b79e-102">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="0b79e-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="0b79e-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="0b79e-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0b79e-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="0b79e-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="0b79e-105">Contidos em</span><span class="sxs-lookup"><span data-stu-id="0b79e-105">Contained in</span></span>

|<span data-ttu-id="0b79e-106">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="0b79e-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="0b79e-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="0b79e-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="0b79e-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="0b79e-108">Description</span></span>](description.md)|
|[<span data-ttu-id="0b79e-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="0b79e-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="0b79e-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="0b79e-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="0b79e-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="0b79e-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="0b79e-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="0b79e-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="0b79e-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="0b79e-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="0b79e-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="0b79e-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="0b79e-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0b79e-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="0b79e-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="0b79e-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="0b79e-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="0b79e-117">Attributes</span></span>

|<span data-ttu-id="0b79e-118">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="0b79e-118">**Attribute**</span></span>|<span data-ttu-id="0b79e-119">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="0b79e-119">**Type**</span></span>|<span data-ttu-id="0b79e-120">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="0b79e-120">**Required**</span></span>|<span data-ttu-id="0b79e-121">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="0b79e-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0b79e-122">Locale</span><span class="sxs-lookup"><span data-stu-id="0b79e-122">Locale</span></span>|<span data-ttu-id="0b79e-123">string</span><span class="sxs-lookup"><span data-stu-id="0b79e-123">string</span></span>|<span data-ttu-id="0b79e-124">obrigatório</span><span class="sxs-lookup"><span data-stu-id="0b79e-124">required</span></span>|<span data-ttu-id="0b79e-125">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="0b79e-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="0b79e-126">Valor</span><span class="sxs-lookup"><span data-stu-id="0b79e-126">Value</span></span>|<span data-ttu-id="0b79e-127">string</span><span class="sxs-lookup"><span data-stu-id="0b79e-127">string</span></span>|<span data-ttu-id="0b79e-128">obrigatório</span><span class="sxs-lookup"><span data-stu-id="0b79e-128">required</span></span>|<span data-ttu-id="0b79e-129">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="0b79e-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="0b79e-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="0b79e-130">See also</span></span>

- [<span data-ttu-id="0b79e-131">Localização para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="0b79e-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
