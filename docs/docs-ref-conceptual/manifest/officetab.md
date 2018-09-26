# <a name="officetab-element"></a><span data-ttu-id="691a1-101">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="691a1-101">OfficeTab element</span></span>

<span data-ttu-id="691a1-p101">Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão (**Início**, **Mensagem** ou **Reunião**) ou uma guia personalizada definida pelo suplemento. Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="691a1-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="691a1-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="691a1-105">Child elements</span></span>

|  <span data-ttu-id="691a1-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="691a1-106">Element</span></span> |  <span data-ttu-id="691a1-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="691a1-107">Required</span></span>  |  <span data-ttu-id="691a1-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="691a1-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="691a1-109">Grupo</span><span class="sxs-lookup"><span data-stu-id="691a1-109">Group</span></span>      | <span data-ttu-id="691a1-110">Sim</span><span class="sxs-lookup"><span data-stu-id="691a1-110">Yes</span></span> |  <span data-ttu-id="691a1-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="691a1-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="691a1-113">A seguir estão os valores válidos de `id` por host.</span><span class="sxs-lookup"><span data-stu-id="691a1-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="691a1-114">Valores em **negrito** são suportados no desktop e online (por exemplo, Word 2016 ou posterior para Windows e o Word Online).</span><span class="sxs-lookup"><span data-stu-id="691a1-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="691a1-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="691a1-115">Outlook</span></span>

- <span data-ttu-id="691a1-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="691a1-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="691a1-117">Word</span><span class="sxs-lookup"><span data-stu-id="691a1-117">Word</span></span>

- <span data-ttu-id="691a1-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="691a1-118">**TabHome**</span></span>
- <span data-ttu-id="691a1-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="691a1-119">**TabInsert**</span></span>
- <span data-ttu-id="691a1-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="691a1-120">TabWordDesign</span></span>
- <span data-ttu-id="691a1-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="691a1-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="691a1-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="691a1-122">TabReferences</span></span>
- <span data-ttu-id="691a1-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="691a1-123">TabMailings</span></span>
- <span data-ttu-id="691a1-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="691a1-124">TabReviewWord</span></span>
- <span data-ttu-id="691a1-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="691a1-125">**TabView**</span></span>
- <span data-ttu-id="691a1-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="691a1-126">TabDeveloper</span></span>
- <span data-ttu-id="691a1-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="691a1-127">TabAddIns</span></span>
- <span data-ttu-id="691a1-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="691a1-128">TabBlogPost</span></span>
- <span data-ttu-id="691a1-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="691a1-129">TabBlogInsert</span></span>
- <span data-ttu-id="691a1-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="691a1-130">TabPrintPreview</span></span>
- <span data-ttu-id="691a1-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="691a1-131">TabOutlining</span></span>
- <span data-ttu-id="691a1-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="691a1-132">TabConflicts</span></span>
- <span data-ttu-id="691a1-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="691a1-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="691a1-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="691a1-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="691a1-135">Excel</span><span class="sxs-lookup"><span data-stu-id="691a1-135">Excel</span></span>

- <span data-ttu-id="691a1-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="691a1-136">**TabHome**</span></span>
- <span data-ttu-id="691a1-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="691a1-137">**TabInsert**</span></span>
- <span data-ttu-id="691a1-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="691a1-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="691a1-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="691a1-139">TabFormulas</span></span>
- <span data-ttu-id="691a1-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="691a1-140">**TabData**</span></span>
- <span data-ttu-id="691a1-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="691a1-141">**TabReview**</span></span>
- <span data-ttu-id="691a1-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="691a1-142">**TabView**</span></span>
- <span data-ttu-id="691a1-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="691a1-143">TabDeveloper</span></span>
- <span data-ttu-id="691a1-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="691a1-144">TabAddIns</span></span>
- <span data-ttu-id="691a1-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="691a1-145">TabPrintPreview</span></span>
- <span data-ttu-id="691a1-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="691a1-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="691a1-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="691a1-147">PowerPoint</span></span>

- <span data-ttu-id="691a1-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="691a1-148">**TabHome**</span></span>
- <span data-ttu-id="691a1-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="691a1-149">**TabInsert**</span></span>
- <span data-ttu-id="691a1-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="691a1-150">**TabDesign**</span></span>
- <span data-ttu-id="691a1-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="691a1-151">**TabTransitions**</span></span>
- <span data-ttu-id="691a1-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="691a1-152">**TabAnimations**</span></span>
- <span data-ttu-id="691a1-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="691a1-153">TabSlideShow</span></span>
- <span data-ttu-id="691a1-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="691a1-154">TabReview</span></span>
- <span data-ttu-id="691a1-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="691a1-155">**TabView**</span></span>
- <span data-ttu-id="691a1-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="691a1-156">TabDeveloper</span></span>
- <span data-ttu-id="691a1-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="691a1-157">TabAddIns</span></span>
- <span data-ttu-id="691a1-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="691a1-158">TabPrintPreview</span></span>
- <span data-ttu-id="691a1-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="691a1-159">TabMerge</span></span>
- <span data-ttu-id="691a1-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="691a1-160">TabGrayscale</span></span>
- <span data-ttu-id="691a1-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="691a1-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="691a1-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="691a1-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="691a1-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="691a1-163">TabSlideMaster</span></span>
- <span data-ttu-id="691a1-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="691a1-164">TabHandoutMaster</span></span>
- <span data-ttu-id="691a1-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="691a1-165">TabNotesMaster</span></span>
- <span data-ttu-id="691a1-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="691a1-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="691a1-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="691a1-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="691a1-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="691a1-168">OneNote</span></span>

- <span data-ttu-id="691a1-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="691a1-169">**TabHome**</span></span>
- <span data-ttu-id="691a1-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="691a1-170">**TabInsert**</span></span>
- <span data-ttu-id="691a1-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="691a1-171">**TabView**</span></span>
- <span data-ttu-id="691a1-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="691a1-172">TabDeveloper</span></span>
- <span data-ttu-id="691a1-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="691a1-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="691a1-174">Grupo</span><span class="sxs-lookup"><span data-stu-id="691a1-174">Group</span></span>

<span data-ttu-id="691a1-p104">Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório, e cada **id** deve ser exclusiva dentro do manifesto. A **id** é uma cadeia de caracteres com, no máximo, 125 caracteres. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="691a1-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="691a1-179">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="691a1-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
