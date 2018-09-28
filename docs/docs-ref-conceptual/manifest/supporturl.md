# <a name="supporturl-element"></a><span data-ttu-id="f3b16-101">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="f3b16-101">SupportUrl element</span></span>

<span data-ttu-id="f3b16-102">Especifica a URL de uma página que fornece informações de suporte para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f3b16-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="f3b16-103">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f3b16-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="f3b16-104">Contidos em</span><span class="sxs-lookup"><span data-stu-id="f3b16-104">Contained in</span></span>

[<span data-ttu-id="f3b16-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f3b16-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f3b16-106">Pode conter</span><span class="sxs-lookup"><span data-stu-id="f3b16-106">Can contain</span></span>

|  <span data-ttu-id="f3b16-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="f3b16-107">Element</span></span> | <span data-ttu-id="f3b16-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f3b16-108">Required</span></span> | <span data-ttu-id="f3b16-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="f3b16-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f3b16-110">Override</span><span class="sxs-lookup"><span data-stu-id="f3b16-110">Override</span></span>](override.md)   | <span data-ttu-id="f3b16-111">Não</span><span class="sxs-lookup"><span data-stu-id="f3b16-111">No</span></span> | <span data-ttu-id="f3b16-112">Especifica a configuração de URLs de localidades adicionais</span><span class="sxs-lookup"><span data-stu-id="f3b16-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="f3b16-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="f3b16-113">Attributes</span></span>

|<span data-ttu-id="f3b16-114">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="f3b16-114">**Attribute**</span></span>|<span data-ttu-id="f3b16-115">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f3b16-115">**Type**</span></span>|<span data-ttu-id="f3b16-116">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="f3b16-116">**Required**</span></span>|<span data-ttu-id="f3b16-117">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="f3b16-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f3b16-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f3b16-118">DefaultValue</span></span>|<span data-ttu-id="f3b16-119">URL</span><span class="sxs-lookup"><span data-stu-id="f3b16-119">URL</span></span>|<span data-ttu-id="f3b16-120">obrigatório</span><span class="sxs-lookup"><span data-stu-id="f3b16-120">required</span></span>|<span data-ttu-id="f3b16-121">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f3b16-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
