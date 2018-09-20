# <a name="outlook-javascript-api-requirement-sets"></a><span data-ttu-id="53253-101">Conjuntos de requisito de API do JavaScript do Outlook</span><span class="sxs-lookup"><span data-stu-id="53253-101">Outlook JavaScript API requirement sets</span></span>

<span data-ttu-id="53253-102">Suplementos do Outlook declarar quais versões de API exigem usando o elemento de [requisitos](/javascript/office/manifest/requirements) no seu [manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="53253-102">Outlook add-ins declare what API versions they require by using the [Requirements](/javascript/office/manifest/requirements) element in their [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests).</span></span> <span data-ttu-id="53253-103">Os suplementos do Outlook sempre incluem um elemento [Set](/javascript/office/manifest/set) com um atributo `Name` definido como `Mailbox` e um atributo `MinVersion` definido como o conjunto de requisitos mínimo de API que dê suporte aos cenários do suplemento.</span><span class="sxs-lookup"><span data-stu-id="53253-103">Outlook add-ins always include a [Set](/javascript/office/manifest/set) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.</span></span>

<span data-ttu-id="53253-104">Por exemplo, o seguinte trecho do manifesto indica um conjunto de requisitos mínimo de 1.1:</span><span class="sxs-lookup"><span data-stu-id="53253-104">For example, the following manifest snippet indicates a minimum requirement set of 1.1:</span></span>

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

<span data-ttu-id="53253-105">Todas as APIs do Outlook pertencem ao [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="53253-105">All Outlook APIs belong to the `Mailbox` [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).</span></span> <span data-ttu-id="53253-106">O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs que lançamos pertence a uma versão superior.</span><span class="sxs-lookup"><span data-stu-id="53253-106">The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set.</span></span> <span data-ttu-id="53253-107">Nem todos os clientes do Outlook serão compatíveis com o conjunto mais recente de APIs, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, será compatível com todas as APIs nesse conjunto.</span><span class="sxs-lookup"><span data-stu-id="53253-107">Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, it supports all of the APIs in that requirement set.</span></span>

<span data-ttu-id="53253-p103">A especificação de uma versão mínima de conjunto de requisitos controla em quais clientes do Outlook o suplemento aparecerá. Se um cliente não oferece suporte para o conjunto de requisitos mínimos, ele não carrega o suplemento. Por exemplo, se for especificada a versão 1.3 do conjunto de requisitos, significa que o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão 1.3.</span><span class="sxs-lookup"><span data-stu-id="53253-p103">Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.</span></span>

## <a name="using-apis-from-later-requirement-sets"></a><span data-ttu-id="53253-111">Usar APIs de conjuntos de requisitos posteriores</span><span class="sxs-lookup"><span data-stu-id="53253-111">Using APIs from later requirement sets</span></span>

<span data-ttu-id="53253-p104">Definir um conjunto de requisitos não limita as APIs disponíveis que o suplemento pode usar. Por exemplo, se o suplemento especificar o conjunto de requisitos 1.1, mas estiver sendo executado em um cliente do Outlook que dá suporte à versão 1.3, o suplemento poderá usar APIs do conjunto de requisitos 1.3\.</span><span class="sxs-lookup"><span data-stu-id="53253-p104">Setting a requirement set does not limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set 1.1, but it is running in an Outlook client which support 1.3, the add-in can use APIs from requirement set 1.3\.</span></span>

<span data-ttu-id="53253-114">Para usar as APIs mais recentes, os desenvolvedores podem apenas verificar sua existência usando a técnica JavaScript padrão.</span><span class="sxs-lookup"><span data-stu-id="53253-114">To use newer APIs, developers can just check for their existence by using standard JavaScript technique</span></span>

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

<span data-ttu-id="53253-115">Essas verificações não são necessárias para APIs que estão presentes na versão do conjunto de requisitos especificada no manifesto.</span><span class="sxs-lookup"><span data-stu-id="53253-115">No such checks are necessary for any APIs which are present in the requirement set version specified in in the manifest.</span></span>

## <a name="choosing-a-minimum-requirement-set"></a><span data-ttu-id="53253-116">Escolher um conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="53253-116">Choosing a minimum requirement set</span></span>

<span data-ttu-id="53253-117">Os desenvolvedores devem usar o conjunto de requisitos mínimos que contém o conjunto essencial de APIs para seu cenário, sem o qual o suplemento não funcionará.</span><span class="sxs-lookup"><span data-stu-id="53253-117">Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.</span></span>

## <a name="clients"></a><span data-ttu-id="53253-118">Clientes</span><span class="sxs-lookup"><span data-stu-id="53253-118">Clients</span></span>

<span data-ttu-id="53253-119">Os clientes a seguir oferecem suporte para suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="53253-119">The following clients support Outlook add-ins.</span></span>

| <span data-ttu-id="53253-120">Cliente</span><span class="sxs-lookup"><span data-stu-id="53253-120">Client</span></span> | <span data-ttu-id="53253-121">Conjuntos de requisitos de API suportados</span><span class="sxs-lookup"><span data-stu-id="53253-121">Supported API requirement sets</span></span> |
| --- | --- |
| <span data-ttu-id="53253-122">Outlook 2016 (Clique para Executar) para Windows</span><span class="sxs-lookup"><span data-stu-id="53253-122">Outlook 2016 (Click-to-Run) for Windows</span></span> | <span data-ttu-id="53253-123">1.1, 1.2, 1.3, 1,4, 1,5, 1.6</span><span class="sxs-lookup"><span data-stu-id="53253-123">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="53253-124">Outlook 2016 (MSI) para Windows</span><span class="sxs-lookup"><span data-stu-id="53253-124">Outlook 2016 (MSI) for Windows</span></span> | <span data-ttu-id="53253-125">1.1, 1.2, 1.3, 1.4</span><span class="sxs-lookup"><span data-stu-id="53253-125">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="53253-126">Outlook 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="53253-126">Outlook 2016 for Mac</span></span> | <span data-ttu-id="53253-127">1.1, 1.2, 1.3, 1,4, 1,5, 1.6</span><span class="sxs-lookup"><span data-stu-id="53253-127">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="53253-128">Outlook 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="53253-128">Outlook 2013 for Windows</span></span> | <span data-ttu-id="53253-129">1.1, 1.2, 1.3, 1.4</span><span class="sxs-lookup"><span data-stu-id="53253-129">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="53253-130">Outlook para iPhone</span><span class="sxs-lookup"><span data-stu-id="53253-130">Outlook for iPhone</span></span> | <span data-ttu-id="53253-131">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="53253-131">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="53253-132">Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="53253-132">Outlook for Android</span></span> | <span data-ttu-id="53253-133">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="53253-133">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="53253-134">Outlook na Web (Office 365 e Outlook.com)</span><span class="sxs-lookup"><span data-stu-id="53253-134">Outlook on the web (Office 365 and Outlook.com)</span></span> | <span data-ttu-id="53253-135">1.1, 1.2, 1.3, 1,4, 1,5, 1.6</span><span class="sxs-lookup"><span data-stu-id="53253-135">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="53253-136">Outlook Web App (Exchange 2013 local)</span><span class="sxs-lookup"><span data-stu-id="53253-136">Outlook Web App (Exchange 2013 On-Premise)</span></span> | <span data-ttu-id="53253-137">1.1</span><span class="sxs-lookup"><span data-stu-id="53253-137">1.1</span></span> |
| <span data-ttu-id="53253-138">Outlook Web App (Exchange 2016 local)</span><span class="sxs-lookup"><span data-stu-id="53253-138">Outlook Web App (Exchange 2016 On-Premise)</span></span> | <span data-ttu-id="53253-p105">1.1, 1.2. 1.3</span><span class="sxs-lookup"><span data-stu-id="53253-p105">1.1, 1.2. 1.3</span></span> |

> [!NOTE] 
> <span data-ttu-id="53253-141">Suporte para 1.3 no Outlook 2013 foi adicionado como parte do [8 de dezembro 2015, atualizar para o Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349).</span><span class="sxs-lookup"><span data-stu-id="53253-141">Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349).</span></span> <span data-ttu-id="53253-142">Suporte para 1,4 no Outlook 2013 foi adicionado como parte da [13 de setembro, 2016, atualizar para o Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).</span><span class="sxs-lookup"><span data-stu-id="53253-142">Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).</span></span>
