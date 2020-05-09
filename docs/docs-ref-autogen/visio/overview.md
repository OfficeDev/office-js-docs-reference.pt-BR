---
title: Referência da API JavaScript do Office
description: Os conjuntos de requisitos de APIs JavaScript do Office por host.
ms.date: 05/05/2020
ms.openlocfilehash: 3a32c47b23fd6635c4c2b44b58ee9b351fffd8d5
ms.sourcegitcommit: 23d9a58660cb1dedf0bc414849a5aec519b419b3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/07/2020
ms.locfileid: "44143927"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="5f9ca-103">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="5f9ca-103">Office JavaScript API reference</span></span>

<span data-ttu-id="5f9ca-104">A API JavaScript para Office permite que você crie aplicativos Web que interajam com os modelos de objeto em aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="5f9ca-105">Use esta seção para saber mais sobre as classes, os métodos e outros tipos disponíveis para a criação de suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="5f9ca-106">Veja a seguir uma lista de conjuntos de requisitos específicos do host (e as APIs comuns de vários hosts).</span><span class="sxs-lookup"><span data-stu-id="5f9ca-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="5f9ca-107">Cada item vincula a uma versão da documentação de referência da API que é suportada por esse conjunto de requisitos (por exemplo, ExcelApi 1,3 mostra APIs no ExcelApi 1,1, 1,2, 1,3, bem como a API comum).</span><span class="sxs-lookup"><span data-stu-id="5f9ca-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="5f9ca-108">`ExcelApiOnline 1.1`é um conjunto de requisitos especiais.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="5f9ca-109">Ele contém as APIs mais recentes para o Excel na Web, mas essas APIs podem não ser totalmente suportadas em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="5f9ca-110">Confira o [conjunto de requisitos da API JavaScript do Excel online somente](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="5f9ca-111">Escolha um link nesta página para exibir a documentação de referência para APIs compatíveis com o conjunto de requisitos especificado ou use o menu suspenso de seleção de filtro acima do Sumário para alterar o conjunto de requisitos a qualquer momento.</span><span class="sxs-lookup"><span data-stu-id="5f9ca-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="5f9ca-112">Excel</span><span class="sxs-lookup"><span data-stu-id="5f9ca-112">Excel</span></span>

- [<span data-ttu-id="5f9ca-113">Visualização do ExcelApi</span><span class="sxs-lookup"><span data-stu-id="5f9ca-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="5f9ca-114">ExcelApiOnline 1,1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="5f9ca-115">ExcelApi 1,11</span><span class="sxs-lookup"><span data-stu-id="5f9ca-115">ExcelApi 1.11</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="5f9ca-116">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="5f9ca-116">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="5f9ca-117">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="5f9ca-117">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="5f9ca-118">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="5f9ca-118">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="5f9ca-119">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="5f9ca-119">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="5f9ca-120">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="5f9ca-120">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="5f9ca-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="5f9ca-121">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="5f9ca-122">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="5f9ca-122">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="5f9ca-123">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="5f9ca-123">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="5f9ca-124">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="5f9ca-124">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="5f9ca-125">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-125">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="5f9ca-126">OneNote</span><span class="sxs-lookup"><span data-stu-id="5f9ca-126">OneNote</span></span>

- [<span data-ttu-id="5f9ca-127">OneNote 1,1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-127">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="5f9ca-128">Outlook</span><span class="sxs-lookup"><span data-stu-id="5f9ca-128">Outlook</span></span>

- [<span data-ttu-id="5f9ca-129">Visualização da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5f9ca-129">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="5f9ca-130">Caixa de correio 1.8</span><span class="sxs-lookup"><span data-stu-id="5f9ca-130">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="5f9ca-131">Caixa de correio 1.7</span><span class="sxs-lookup"><span data-stu-id="5f9ca-131">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="5f9ca-132">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="5f9ca-132">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="5f9ca-133"> Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="5f9ca-133">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="5f9ca-134"> Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="5f9ca-134">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="5f9ca-135"> Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="5f9ca-135">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="5f9ca-136">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="5f9ca-136">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="5f9ca-137"> Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-137">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="5f9ca-138">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5f9ca-138">PowerPoint</span></span>

- [<span data-ttu-id="5f9ca-139">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-139">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="5f9ca-140">Visio</span><span class="sxs-lookup"><span data-stu-id="5f9ca-140">Visio</span></span>

- [<span data-ttu-id="5f9ca-141">VisioApi 1,1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-141">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="5f9ca-142">Word</span><span class="sxs-lookup"><span data-stu-id="5f9ca-142">Word</span></span>

- [<span data-ttu-id="5f9ca-143">Visualização do Word</span><span class="sxs-lookup"><span data-stu-id="5f9ca-143">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="5f9ca-144">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="5f9ca-144">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="5f9ca-145">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="5f9ca-145">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="5f9ca-146">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5f9ca-146">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="5f9ca-147">API Comum</span><span class="sxs-lookup"><span data-stu-id="5f9ca-147">Common API</span></span>

- [<span data-ttu-id="5f9ca-148">API Comum</span><span class="sxs-lookup"><span data-stu-id="5f9ca-148">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="5f9ca-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="5f9ca-149">See also</span></span>

- [<span data-ttu-id="5f9ca-150">Sobre os Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5f9ca-150">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="5f9ca-151">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5f9ca-151">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="5f9ca-152">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="5f9ca-152">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- <span data-ttu-id="5f9ca-153">[Explore a API JavaScript do Office usando o Script Lab](/office/dev/add-ins/overview/explore-with-script-lab).</span><span class="sxs-lookup"><span data-stu-id="5f9ca-153">[Explore Office JavaScript API using Script Lab](/office/dev/add-ins/overview/explore-with-script-lab)</span></span>
