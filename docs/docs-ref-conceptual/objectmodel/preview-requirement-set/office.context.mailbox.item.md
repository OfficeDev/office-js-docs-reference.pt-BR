
# <a name="item"></a><span data-ttu-id="4e51e-101">item</span><span class="sxs-lookup"><span data-stu-id="4e51e-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4e51e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4e51e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4e51e-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4e51e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-105">Requirements</span></span>

|<span data-ttu-id="4e51e-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-106">Requirement</span></span>|<span data-ttu-id="4e51e-107">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-109">1.0</span></span>|
|[<span data-ttu-id="4e51e-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="4e51e-111">Restricted</span></span>|
|[<span data-ttu-id="4e51e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4e51e-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4e51e-114">Members and methods</span></span>

| <span data-ttu-id="4e51e-115">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-115">Member</span></span> | <span data-ttu-id="4e51e-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4e51e-117">attachments</span><span class="sxs-lookup"><span data-stu-id="4e51e-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="4e51e-118">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-118">Member</span></span> |
| [<span data-ttu-id="4e51e-119">bcc</span><span class="sxs-lookup"><span data-stu-id="4e51e-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4e51e-120">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-120">Member</span></span> |
| [<span data-ttu-id="4e51e-121">body</span><span class="sxs-lookup"><span data-stu-id="4e51e-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="4e51e-122">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-122">Member</span></span> |
| [<span data-ttu-id="4e51e-123">cc</span><span class="sxs-lookup"><span data-stu-id="4e51e-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4e51e-124">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-124">Member</span></span> |
| [<span data-ttu-id="4e51e-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="4e51e-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4e51e-126">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-126">Member</span></span> |
| [<span data-ttu-id="4e51e-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4e51e-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4e51e-128">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-128">Member</span></span> |
| [<span data-ttu-id="4e51e-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4e51e-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4e51e-130">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-130">Member</span></span> |
| [<span data-ttu-id="4e51e-131">end</span><span class="sxs-lookup"><span data-stu-id="4e51e-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="4e51e-132">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-132">Member</span></span> |
| [<span data-ttu-id="4e51e-133">from</span><span class="sxs-lookup"><span data-stu-id="4e51e-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="4e51e-134">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-134">Member</span></span> |
| [<span data-ttu-id="4e51e-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4e51e-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4e51e-136">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-136">Member</span></span> |
| [<span data-ttu-id="4e51e-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="4e51e-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4e51e-138">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-138">Member</span></span> |
| [<span data-ttu-id="4e51e-139">itemId</span><span class="sxs-lookup"><span data-stu-id="4e51e-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4e51e-140">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-140">Member</span></span> |
| [<span data-ttu-id="4e51e-141">itemType</span><span class="sxs-lookup"><span data-stu-id="4e51e-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="4e51e-142">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-142">Member</span></span> |
| [<span data-ttu-id="4e51e-143">location</span><span class="sxs-lookup"><span data-stu-id="4e51e-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="4e51e-144">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-144">Member</span></span> |
| [<span data-ttu-id="4e51e-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4e51e-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4e51e-146">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-146">Member</span></span> |
| [<span data-ttu-id="4e51e-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4e51e-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="4e51e-148">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-148">Member</span></span> |
| [<span data-ttu-id="4e51e-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4e51e-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4e51e-150">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-150">Member</span></span> |
| [<span data-ttu-id="4e51e-151">organizer</span><span class="sxs-lookup"><span data-stu-id="4e51e-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="4e51e-152">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-152">Member</span></span> |
| [<span data-ttu-id="4e51e-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="4e51e-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="4e51e-154">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-154">Member</span></span> |
| [<span data-ttu-id="4e51e-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4e51e-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4e51e-156">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-156">Member</span></span> |
| [<span data-ttu-id="4e51e-157">sender</span><span class="sxs-lookup"><span data-stu-id="4e51e-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="4e51e-158">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-158">Member</span></span> |
| [<span data-ttu-id="4e51e-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="4e51e-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4e51e-160">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-160">Member</span></span> |
| [<span data-ttu-id="4e51e-161">start</span><span class="sxs-lookup"><span data-stu-id="4e51e-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="4e51e-162">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-162">Member</span></span> |
| [<span data-ttu-id="4e51e-163">subject</span><span class="sxs-lookup"><span data-stu-id="4e51e-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="4e51e-164">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-164">Member</span></span> |
| [<span data-ttu-id="4e51e-165">to</span><span class="sxs-lookup"><span data-stu-id="4e51e-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="4e51e-166">Membro</span><span class="sxs-lookup"><span data-stu-id="4e51e-166">Member</span></span> |
| [<span data-ttu-id="4e51e-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4e51e-168">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-168">Method</span></span> |
| [<span data-ttu-id="4e51e-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="4e51e-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="4e51e-170">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-170">Method</span></span> |
| [<span data-ttu-id="4e51e-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4e51e-172">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-172">Method</span></span> |
| [<span data-ttu-id="4e51e-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4e51e-174">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-174">Method</span></span> |
| [<span data-ttu-id="4e51e-175">close</span><span class="sxs-lookup"><span data-stu-id="4e51e-175">close</span></span>](#close) | <span data-ttu-id="4e51e-176">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-176">Method</span></span> |
| [<span data-ttu-id="4e51e-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4e51e-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="4e51e-178">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-178">Method</span></span> |
| [<span data-ttu-id="4e51e-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4e51e-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="4e51e-180">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-180">Method</span></span> |
| [<span data-ttu-id="4e51e-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="4e51e-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="4e51e-182">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-182">Method</span></span> |
| [<span data-ttu-id="4e51e-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4e51e-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="4e51e-184">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-184">Method</span></span> |
| [<span data-ttu-id="4e51e-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4e51e-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="4e51e-186">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-186">Method</span></span> |
| [<span data-ttu-id="4e51e-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="4e51e-188">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-188">Method</span></span> |
| [<span data-ttu-id="4e51e-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4e51e-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4e51e-190">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-190">Method</span></span> |
| [<span data-ttu-id="4e51e-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4e51e-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4e51e-192">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-192">Method</span></span> |
| [<span data-ttu-id="4e51e-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4e51e-194">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-194">Method</span></span> |
| [<span data-ttu-id="4e51e-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="4e51e-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="4e51e-196">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-196">Method</span></span> |
| [<span data-ttu-id="4e51e-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4e51e-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4e51e-198">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-198">Method</span></span> |
| [<span data-ttu-id="4e51e-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4e51e-200">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-200">Method</span></span> |
| [<span data-ttu-id="4e51e-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4e51e-202">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-202">Method</span></span> |
| [<span data-ttu-id="4e51e-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4e51e-204">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-204">Method</span></span> |
| [<span data-ttu-id="4e51e-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4e51e-206">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-206">Method</span></span> |
| [<span data-ttu-id="4e51e-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4e51e-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4e51e-208">Método</span><span class="sxs-lookup"><span data-stu-id="4e51e-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4e51e-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-209">Example</span></span>

<span data-ttu-id="4e51e-210">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="4e51e-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="4e51e-211">Membros</span><span class="sxs-lookup"><span data-stu-id="4e51e-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="4e51e-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4e51e-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="4e51e-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-215">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="4e51e-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4e51e-216">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4e51e-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-217">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-217">Type:</span></span>

*   <span data-ttu-id="4e51e-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4e51e-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-219">Requirements</span></span>

|<span data-ttu-id="4e51e-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-220">Requirement</span></span>|<span data-ttu-id="4e51e-221">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-223">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-223">1.0</span></span>|
|[<span data-ttu-id="4e51e-224">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-225">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-227">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-228">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-228">Example</span></span>

<span data-ttu-id="4e51e-229">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4e51e-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4e51e-231">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4e51e-232">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="4e51e-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-233">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-233">Type:</span></span>

*   [<span data-ttu-id="4e51e-234">Destinatários</span><span class="sxs-lookup"><span data-stu-id="4e51e-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4e51e-235">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-235">Requirements</span></span>

|<span data-ttu-id="4e51e-236">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-236">Requirement</span></span>|<span data-ttu-id="4e51e-237">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-238">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-239">1.1</span><span class="sxs-lookup"><span data-stu-id="4e51e-239">1.1</span></span>|
|[<span data-ttu-id="4e51e-240">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-241">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-243">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-244">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="4e51e-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="4e51e-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="4e51e-246">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-247">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-247">Type:</span></span>

*   [<span data-ttu-id="4e51e-248">Corpo</span><span class="sxs-lookup"><span data-stu-id="4e51e-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="4e51e-249">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-249">Requirements</span></span>

|<span data-ttu-id="4e51e-250">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-250">Requirement</span></span>|<span data-ttu-id="4e51e-251">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-252">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-253">1.1</span><span class="sxs-lookup"><span data-stu-id="4e51e-253">1.1</span></span>|
|[<span data-ttu-id="4e51e-254">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-255">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-256">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-257">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4e51e-258">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4e51e-259">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4e51e-260">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-261">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-261">Read mode</span></span>

<span data-ttu-id="4e51e-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-264">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-264">Compose mode</span></span>

<span data-ttu-id="4e51e-265">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-266">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-266">Type:</span></span>

*   <span data-ttu-id="4e51e-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-268">Requirements</span></span>

|<span data-ttu-id="4e51e-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-269">Requirement</span></span>|<span data-ttu-id="4e51e-270">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-272">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-272">1.0</span></span>|
|[<span data-ttu-id="4e51e-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-274">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-276">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4e51e-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4e51e-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="4e51e-279">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="4e51e-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4e51e-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4e51e-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-284">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-284">Type:</span></span>

*   <span data-ttu-id="4e51e-285">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-286">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-286">Requirements</span></span>

|<span data-ttu-id="4e51e-287">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-287">Requirement</span></span>|<span data-ttu-id="4e51e-288">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-289">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-290">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-290">1.0</span></span>|
|[<span data-ttu-id="4e51e-291">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-292">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-293">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-294">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="4e51e-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4e51e-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="4e51e-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-298">Type:</span></span>

*   <span data-ttu-id="4e51e-299">Data</span><span class="sxs-lookup"><span data-stu-id="4e51e-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-300">Requirements</span></span>

|<span data-ttu-id="4e51e-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-301">Requirement</span></span>|<span data-ttu-id="4e51e-302">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-304">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-304">1.0</span></span>|
|[<span data-ttu-id="4e51e-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-306">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-308">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4e51e-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4e51e-310">dateTimeModified :Date</span></span>

<span data-ttu-id="4e51e-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-313">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-314">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-314">Type:</span></span>

*   <span data-ttu-id="4e51e-315">Data</span><span class="sxs-lookup"><span data-stu-id="4e51e-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-316">Requirements</span></span>

|<span data-ttu-id="4e51e-317">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-317">Requirement</span></span>|<span data-ttu-id="4e51e-318">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-320">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-320">1.0</span></span>|
|[<span data-ttu-id="4e51e-321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-322">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-324">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4e51e-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4e51e-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4e51e-327">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="4e51e-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4e51e-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-330">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-330">Read mode</span></span>

<span data-ttu-id="4e51e-331">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-332">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-332">Compose mode</span></span>

<span data-ttu-id="4e51e-333">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4e51e-334">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-335">Type:</span></span>

*   <span data-ttu-id="4e51e-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4e51e-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-337">Requirements</span></span>

|<span data-ttu-id="4e51e-338">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-338">Requirement</span></span>|<span data-ttu-id="4e51e-339">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-340">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-341">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-341">1.0</span></span>|
|[<span data-ttu-id="4e51e-342">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-343">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-344">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-345">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-346">Example</span></span>

<span data-ttu-id="4e51e-347">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="4e51e-348">de:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4e51e-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="4e51e-349">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4e51e-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-352">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-353">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-353">Read mode</span></span>

<span data-ttu-id="4e51e-354">O `from` propriedade retornará uma `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="4e51e-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-355">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-355">Compose mode</span></span>

<span data-ttu-id="4e51e-356">O `from` propriedade retornará uma `From` objeto que fornece um método para obter o de valor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e51e-357">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-357">Type:</span></span>

*   <span data-ttu-id="4e51e-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4e51e-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-359">Requirements</span></span>

|<span data-ttu-id="4e51e-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4e51e-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-362">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-362">1.0</span></span>|<span data-ttu-id="4e51e-363">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-363">Preview</span></span>|
|[<span data-ttu-id="4e51e-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-365">ReadItem</span></span>|<span data-ttu-id="4e51e-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-367">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-368">Read</span><span class="sxs-lookup"><span data-stu-id="4e51e-368">Read</span></span>|<span data-ttu-id="4e51e-369">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="4e51e-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4e51e-370">internetMessageId :String</span></span>

<span data-ttu-id="4e51e-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-373">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-373">Type:</span></span>

*   <span data-ttu-id="4e51e-374">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-375">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-375">Requirements</span></span>

|<span data-ttu-id="4e51e-376">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-376">Requirement</span></span>|<span data-ttu-id="4e51e-377">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-378">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-379">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-379">1.0</span></span>|
|[<span data-ttu-id="4e51e-380">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-381">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-382">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-383">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-384">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4e51e-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4e51e-385">itemClass :String</span></span>

<span data-ttu-id="4e51e-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4e51e-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4e51e-390">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-390">Type</span></span>|<span data-ttu-id="4e51e-391">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-391">Description</span></span>|<span data-ttu-id="4e51e-392">classe de item</span><span class="sxs-lookup"><span data-stu-id="4e51e-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4e51e-393">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="4e51e-393">Appointment items</span></span>|<span data-ttu-id="4e51e-394">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="4e51e-395">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="4e51e-395">Message items</span></span>|<span data-ttu-id="4e51e-396">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="4e51e-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4e51e-397">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-398">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-398">Type:</span></span>

*   <span data-ttu-id="4e51e-399">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-400">Requirements</span></span>

|<span data-ttu-id="4e51e-401">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-401">Requirement</span></span>|<span data-ttu-id="4e51e-402">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-403">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-404">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-404">1.0</span></span>|
|[<span data-ttu-id="4e51e-405">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-406">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-407">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-408">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-409">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4e51e-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4e51e-410">(nullable) itemId :String</span></span>

<span data-ttu-id="4e51e-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-413">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e51e-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4e51e-414">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4e51e-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4e51e-415">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4e51e-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4e51e-416">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4e51e-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4e51e-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-419">Type:</span></span>

*   <span data-ttu-id="4e51e-420">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-421">Requirements</span></span>

|<span data-ttu-id="4e51e-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-422">Requirement</span></span>|<span data-ttu-id="4e51e-423">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-425">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-425">1.0</span></span>|
|[<span data-ttu-id="4e51e-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-427">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-429">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-430">Example</span></span>

<span data-ttu-id="4e51e-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="4e51e-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4e51e-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4e51e-434">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="4e51e-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4e51e-435">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-436">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-436">Type:</span></span>

*   [<span data-ttu-id="4e51e-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4e51e-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4e51e-438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-438">Requirements</span></span>

|<span data-ttu-id="4e51e-439">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-439">Requirement</span></span>|<span data-ttu-id="4e51e-440">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-441">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-442">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-442">1.0</span></span>|
|[<span data-ttu-id="4e51e-443">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-444">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-445">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-446">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-447">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="4e51e-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4e51e-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="4e51e-449">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-450">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-450">Read mode</span></span>

<span data-ttu-id="4e51e-451">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-452">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-452">Compose mode</span></span>

<span data-ttu-id="4e51e-453">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-454">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-454">Type:</span></span>

*   <span data-ttu-id="4e51e-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4e51e-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-456">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-456">Requirements</span></span>

|<span data-ttu-id="4e51e-457">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-457">Requirement</span></span>|<span data-ttu-id="4e51e-458">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-459">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-460">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-460">1.0</span></span>|
|[<span data-ttu-id="4e51e-461">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-462">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-463">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-464">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-465">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4e51e-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4e51e-466">normalizedSubject :String</span></span>

<span data-ttu-id="4e51e-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4e51e-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="4e51e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-471">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-471">Type:</span></span>

*   <span data-ttu-id="4e51e-472">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-473">Requirements</span></span>

|<span data-ttu-id="4e51e-474">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-474">Requirement</span></span>|<span data-ttu-id="4e51e-475">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-477">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-477">1.0</span></span>|
|[<span data-ttu-id="4e51e-478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-479">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-480">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-481">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="4e51e-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4e51e-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="4e51e-484">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-485">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-485">Type:</span></span>

*   [<span data-ttu-id="4e51e-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4e51e-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4e51e-487">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-487">Requirements</span></span>

|<span data-ttu-id="4e51e-488">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-488">Requirement</span></span>|<span data-ttu-id="4e51e-489">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-490">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-491">1.3</span><span class="sxs-lookup"><span data-stu-id="4e51e-491">1.3</span></span>|
|[<span data-ttu-id="4e51e-492">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-493">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-494">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-495">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4e51e-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4e51e-497">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="4e51e-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4e51e-498">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-499">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-499">Read mode</span></span>

<span data-ttu-id="4e51e-500">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-501">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-501">Compose mode</span></span>

<span data-ttu-id="4e51e-502">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-503">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-503">Type:</span></span>

*   <span data-ttu-id="4e51e-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-505">Requirements</span></span>

|<span data-ttu-id="4e51e-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-506">Requirement</span></span>|<span data-ttu-id="4e51e-507">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-509">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-509">1.0</span></span>|
|[<span data-ttu-id="4e51e-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-511">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-513">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="4e51e-515">Organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4e51e-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="4e51e-516">Obtém o endereço de email do organizador da reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-517">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-517">Read mode</span></span>

<span data-ttu-id="4e51e-518">O `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-519">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-519">Compose mode</span></span>

<span data-ttu-id="4e51e-520">O `organizer` propriedade retorna um objeto do [Organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="4e51e-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-521">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-521">Type:</span></span>

*   <span data-ttu-id="4e51e-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4e51e-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-523">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-523">Requirements</span></span>

|<span data-ttu-id="4e51e-524">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4e51e-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-526">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-526">1.0</span></span>|<span data-ttu-id="4e51e-527">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-527">Preview</span></span>|
|[<span data-ttu-id="4e51e-528">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-529">ReadItem</span></span>|<span data-ttu-id="4e51e-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-531">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-532">Read</span><span class="sxs-lookup"><span data-stu-id="4e51e-532">Read</span></span>|<span data-ttu-id="4e51e-533">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-534">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="4e51e-535">Recorrência (anulável):[Recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="4e51e-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="4e51e-536">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4e51e-537">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4e51e-538">Leitura e redação modos para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4e51e-539">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="4e51e-540">O `recurrence` propriedade retorna um objeto de [Recorrência](/javascript/api/outlook/office.recurrence) para solicitações de reuniões ou compromissos recorrentes, se um item é uma série ou uma instância de uma série.</span><span class="sxs-lookup"><span data-stu-id="4e51e-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4e51e-541">`null`é retornada para um único compromissos e solicitações de reunião de compromissos de única.</span><span class="sxs-lookup"><span data-stu-id="4e51e-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4e51e-542">`undefined`é retornado para mensagens que não fazem solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4e51e-543">Observação: Solicitações de reunião tem um `itemClass` valor IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="4e51e-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4e51e-544">Observação: Se o objeto de recorrência for `null`, isto indica que o objeto é um compromisso único ou uma solicitação de reunião de um único compromisso e não fizer parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="4e51e-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-545">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-545">Type:</span></span>

* [<span data-ttu-id="4e51e-546">Recorrência</span><span class="sxs-lookup"><span data-stu-id="4e51e-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="4e51e-547">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-547">Requirement</span></span>|<span data-ttu-id="4e51e-548">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-549">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-550">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-550">Preview</span></span>|
|[<span data-ttu-id="4e51e-551">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-552">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-553">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-554">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4e51e-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4e51e-556">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="4e51e-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4e51e-557">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-558">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-558">Read mode</span></span>

<span data-ttu-id="4e51e-559">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-560">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-560">Compose mode</span></span>

<span data-ttu-id="4e51e-561">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-562">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-562">Type:</span></span>

*   <span data-ttu-id="4e51e-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-564">Requirements</span></span>

|<span data-ttu-id="4e51e-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-565">Requirement</span></span>|<span data-ttu-id="4e51e-566">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-568">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-568">1.0</span></span>|
|[<span data-ttu-id="4e51e-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-570">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-571">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-572">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-573">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="4e51e-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4e51e-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="4e51e-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4e51e-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-579">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `sender` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-580">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-580">Type:</span></span>

*   [<span data-ttu-id="4e51e-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4e51e-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4e51e-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-582">Requirements</span></span>

|<span data-ttu-id="4e51e-583">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-583">Requirement</span></span>|<span data-ttu-id="4e51e-584">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-585">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-586">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-586">1.0</span></span>|
|[<span data-ttu-id="4e51e-587">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-588">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-589">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-590">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-591">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4e51e-592">seriesId (anulável): cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="4e51e-593">Obtém a identificação da série que uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="4e51e-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4e51e-594">No OWA e no Outlook, o `seriesId` retornará a identificação de serviços Web do Exchange (EWS) do item pai (série) que este item pertence.</span><span class="sxs-lookup"><span data-stu-id="4e51e-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4e51e-595">No entanto, em iOS e Android, o `seriesId` retornará a identificação do restante do item pai.</span><span class="sxs-lookup"><span data-stu-id="4e51e-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-596">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e51e-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4e51e-597">O `seriesId` propriedade não é idêntica às IDs do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4e51e-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4e51e-598">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4e51e-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4e51e-599">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4e51e-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4e51e-600">O `seriesId` propriedade retornará `null` para itens que não têm itens pai como único compromissos, itens de série, ou solicitações de reunião e retorna `undefined` para todos os itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="4e51e-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-601">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-601">Type:</span></span>

* <span data-ttu-id="4e51e-602">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-603">Requirements</span></span>

|<span data-ttu-id="4e51e-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-604">Requirement</span></span>|<span data-ttu-id="4e51e-605">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-607">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-607">Preview</span></span>|
|[<span data-ttu-id="4e51e-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-609">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-611">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-612">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4e51e-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4e51e-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4e51e-614">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="4e51e-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4e51e-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-617">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-617">Read mode</span></span>

<span data-ttu-id="4e51e-618">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-619">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-619">Compose mode</span></span>

<span data-ttu-id="4e51e-620">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4e51e-621">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-622">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-622">Type:</span></span>

*   <span data-ttu-id="4e51e-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4e51e-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-624">Requirements</span></span>

|<span data-ttu-id="4e51e-625">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-625">Requirement</span></span>|<span data-ttu-id="4e51e-626">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-627">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-628">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-628">1.0</span></span>|
|[<span data-ttu-id="4e51e-629">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-630">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-631">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-632">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-633">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-633">Example</span></span>

<span data-ttu-id="4e51e-634">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="4e51e-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4e51e-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="4e51e-636">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4e51e-637">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="4e51e-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-638">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-638">Read mode</span></span>

<span data-ttu-id="4e51e-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-641">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-641">Compose mode</span></span>

<span data-ttu-id="4e51e-642">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="4e51e-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e51e-643">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-643">Type:</span></span>

*   <span data-ttu-id="4e51e-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4e51e-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-645">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-645">Requirements</span></span>

|<span data-ttu-id="4e51e-646">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-646">Requirement</span></span>|<span data-ttu-id="4e51e-647">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-648">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-649">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-649">1.0</span></span>|
|[<span data-ttu-id="4e51e-650">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-651">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-652">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-653">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4e51e-654">para: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4e51e-655">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4e51e-656">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e51e-657">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-657">Read mode</span></span>

<span data-ttu-id="4e51e-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4e51e-660">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4e51e-660">Compose mode</span></span>

<span data-ttu-id="4e51e-661">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4e51e-662">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4e51e-662">Type:</span></span>

*   <span data-ttu-id="4e51e-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4e51e-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-664">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-664">Requirements</span></span>

|<span data-ttu-id="4e51e-665">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-665">Requirement</span></span>|<span data-ttu-id="4e51e-666">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-667">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-668">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-668">1.0</span></span>|
|[<span data-ttu-id="4e51e-669">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-670">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-671">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-672">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-673">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="4e51e-674">Métodos</span><span class="sxs-lookup"><span data-stu-id="4e51e-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4e51e-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e51e-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4e51e-676">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4e51e-677">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="4e51e-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4e51e-678">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4e51e-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-679">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-679">Parameters:</span></span>
|<span data-ttu-id="4e51e-680">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-680">Name</span></span>|<span data-ttu-id="4e51e-681">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-681">Type</span></span>|<span data-ttu-id="4e51e-682">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-682">Attributes</span></span>|<span data-ttu-id="4e51e-683">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4e51e-684">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-684">String</span></span>||<span data-ttu-id="4e51e-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4e51e-687">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-687">String</span></span>||<span data-ttu-id="4e51e-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4e51e-690">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-690">Object</span></span>|<span data-ttu-id="4e51e-691">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-691">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-692">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-693">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-693">Object</span></span>|<span data-ttu-id="4e51e-694">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-694">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-695">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4e51e-696">Booliano</span><span class="sxs-lookup"><span data-stu-id="4e51e-696">Boolean</span></span>|<span data-ttu-id="4e51e-697">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-697">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-698">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4e51e-699">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-699">function</span></span>|<span data-ttu-id="4e51e-700">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-700">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-701">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e51e-702">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4e51e-703">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e51e-704">Erros</span><span class="sxs-lookup"><span data-stu-id="4e51e-704">Errors</span></span>

|<span data-ttu-id="4e51e-705">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4e51e-705">Error code</span></span>|<span data-ttu-id="4e51e-706">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4e51e-707">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="4e51e-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4e51e-708">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="4e51e-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4e51e-709">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-710">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-710">Requirements</span></span>

|<span data-ttu-id="4e51e-711">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-711">Requirement</span></span>|<span data-ttu-id="4e51e-712">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-713">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-714">1.1</span><span class="sxs-lookup"><span data-stu-id="4e51e-714">1.1</span></span>|
|[<span data-ttu-id="4e51e-715">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-717">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-718">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e51e-719">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4e51e-719">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="4e51e-720">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="4e51e-721">addFileAttachmentFromBase64Async (base64File, attachmentName, [Opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="4e51e-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4e51e-722">Adiciona um arquivo a partir do base64 codificação para uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4e51e-723">O `addFileAttachmentFromBase64Async` método carrega o arquivo a partir da codificação base64 e o anexa ao item no formato redação.</span><span class="sxs-lookup"><span data-stu-id="4e51e-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="4e51e-724">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="4e51e-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="4e51e-725">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4e51e-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-726">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-726">Parameters:</span></span>
|<span data-ttu-id="4e51e-727">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-727">Name</span></span>|<span data-ttu-id="4e51e-728">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-728">Type</span></span>|<span data-ttu-id="4e51e-729">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-729">Attributes</span></span>|<span data-ttu-id="4e51e-730">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="4e51e-731">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-731">String</span></span>||<span data-ttu-id="4e51e-732">O base64 codificado conteúdo de uma imagem ou um arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="4e51e-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="4e51e-733">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-733">String</span></span>||<span data-ttu-id="4e51e-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4e51e-736">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-736">Object</span></span>|<span data-ttu-id="4e51e-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-737">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-738">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-739">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-739">Object</span></span>|<span data-ttu-id="4e51e-740">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-740">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-741">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4e51e-742">Booliano</span><span class="sxs-lookup"><span data-stu-id="4e51e-742">Boolean</span></span>|<span data-ttu-id="4e51e-743">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-743">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-744">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4e51e-745">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-745">function</span></span>|<span data-ttu-id="4e51e-746">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-746">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-747">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e51e-748">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4e51e-749">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e51e-750">Erros</span><span class="sxs-lookup"><span data-stu-id="4e51e-750">Errors</span></span>

|<span data-ttu-id="4e51e-751">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4e51e-751">Error code</span></span>|<span data-ttu-id="4e51e-752">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4e51e-753">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="4e51e-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4e51e-754">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="4e51e-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4e51e-755">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-756">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-756">Requirements</span></span>

|<span data-ttu-id="4e51e-757">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-757">Requirement</span></span>|<span data-ttu-id="4e51e-758">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-759">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-760">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-760">Preview</span></span>|
|[<span data-ttu-id="4e51e-761">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-763">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-764">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e51e-765">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4e51e-765">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4e51e-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e51e-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4e51e-767">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4e51e-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4e51e-768">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="4e51e-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-769">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-769">Parameters:</span></span>

| <span data-ttu-id="4e51e-770">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-770">Name</span></span> | <span data-ttu-id="4e51e-771">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-771">Type</span></span> | <span data-ttu-id="4e51e-772">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-772">Attributes</span></span> | <span data-ttu-id="4e51e-773">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4e51e-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4e51e-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4e51e-775">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4e51e-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4e51e-776">Função</span><span class="sxs-lookup"><span data-stu-id="4e51e-776">Function</span></span> || <span data-ttu-id="4e51e-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4e51e-780">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-780">Object</span></span> | <span data-ttu-id="4e51e-781">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-781">&lt;optional&gt;</span></span> | <span data-ttu-id="4e51e-782">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4e51e-783">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-783">Object</span></span> | <span data-ttu-id="4e51e-784">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-784">&lt;optional&gt;</span></span> | <span data-ttu-id="4e51e-785">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4e51e-786">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-786">function</span></span>| <span data-ttu-id="4e51e-787">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-787">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-788">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-789">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-789">Requirements</span></span>

|<span data-ttu-id="4e51e-790">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-790">Requirement</span></span>| <span data-ttu-id="4e51e-791">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-792">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e51e-793">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-793">Preview</span></span> |
|[<span data-ttu-id="4e51e-794">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e51e-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-795">ReadItem</span></span> |
|[<span data-ttu-id="4e51e-796">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e51e-797">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="4e51e-798">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4e51e-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e51e-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4e51e-800">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4e51e-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4e51e-804">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4e51e-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4e51e-805">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-806">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-806">Parameters:</span></span>

|<span data-ttu-id="4e51e-807">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-807">Name</span></span>|<span data-ttu-id="4e51e-808">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-808">Type</span></span>|<span data-ttu-id="4e51e-809">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-809">Attributes</span></span>|<span data-ttu-id="4e51e-810">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4e51e-811">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-811">String</span></span>||<span data-ttu-id="4e51e-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4e51e-814">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-814">String</span></span>||<span data-ttu-id="4e51e-p141">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4e51e-817">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-817">Object</span></span>|<span data-ttu-id="4e51e-818">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-818">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-819">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-820">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-820">Object</span></span>|<span data-ttu-id="4e51e-821">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-821">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-822">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4e51e-823">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-823">function</span></span>|<span data-ttu-id="4e51e-824">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-824">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-825">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e51e-826">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4e51e-827">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e51e-828">Erros</span><span class="sxs-lookup"><span data-stu-id="4e51e-828">Errors</span></span>

|<span data-ttu-id="4e51e-829">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4e51e-829">Error code</span></span>|<span data-ttu-id="4e51e-830">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4e51e-831">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-832">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-832">Requirements</span></span>

|<span data-ttu-id="4e51e-833">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-833">Requirement</span></span>|<span data-ttu-id="4e51e-834">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-835">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-836">1.1</span><span class="sxs-lookup"><span data-stu-id="4e51e-836">1.1</span></span>|
|[<span data-ttu-id="4e51e-837">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-839">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-840">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-841">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-841">Example</span></span>

<span data-ttu-id="4e51e-842">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="4e51e-843">close()</span><span class="sxs-lookup"><span data-stu-id="4e51e-843">close()</span></span>

<span data-ttu-id="4e51e-844">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="4e51e-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4e51e-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-847">No Outlook na web, se o item é um compromisso e ele tenha sido salva anteriormente usando `saveAsync`, o usuário é solicitado a salvar, descartar ou Cancelar, mesmo se nenhuma alteração ocorreram desde que o item foi último salvo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4e51e-848">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="4e51e-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-849">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-849">Requirements</span></span>

|<span data-ttu-id="4e51e-850">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-850">Requirement</span></span>|<span data-ttu-id="4e51e-851">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-852">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-853">1.3</span><span class="sxs-lookup"><span data-stu-id="4e51e-853">1.3</span></span>|
|[<span data-ttu-id="4e51e-854">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-855">Restrito</span><span class="sxs-lookup"><span data-stu-id="4e51e-855">Restricted</span></span>|
|[<span data-ttu-id="4e51e-856">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-857">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="4e51e-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4e51e-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="4e51e-859">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-860">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-861">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="4e51e-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4e51e-862">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4e51e-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4e51e-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-866">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-866">Parameters:</span></span>

|<span data-ttu-id="4e51e-867">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-867">Name</span></span>|<span data-ttu-id="4e51e-868">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-868">Type</span></span>|<span data-ttu-id="4e51e-869">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-869">Attributes</span></span>|<span data-ttu-id="4e51e-870">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4e51e-871">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4e51e-871">String &#124; Object</span></span>||<span data-ttu-id="4e51e-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4e51e-874">**OU**</span><span class="sxs-lookup"><span data-stu-id="4e51e-874">**OR**</span></span><br/><span data-ttu-id="4e51e-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4e51e-877">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-877">String</span></span>|<span data-ttu-id="4e51e-878">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-878">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4e51e-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4e51e-882">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-882">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-883">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4e51e-884">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-884">String</span></span>||<span data-ttu-id="4e51e-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4e51e-887">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-887">String</span></span>||<span data-ttu-id="4e51e-888">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="4e51e-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4e51e-889">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-889">String</span></span>||<span data-ttu-id="4e51e-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4e51e-892">Boolean</span><span class="sxs-lookup"><span data-stu-id="4e51e-892">Boolean</span></span>||<span data-ttu-id="4e51e-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4e51e-895">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-895">String</span></span>||<span data-ttu-id="4e51e-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4e51e-899">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-899">function</span></span>|<span data-ttu-id="4e51e-900">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-900">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-901">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-902">Requirements</span></span>

|<span data-ttu-id="4e51e-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-903">Requirement</span></span>|<span data-ttu-id="4e51e-904">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-906">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-906">1.0</span></span>|
|[<span data-ttu-id="4e51e-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-908">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-910">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e51e-911">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4e51e-911">Examples</span></span>

<span data-ttu-id="4e51e-912">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4e51e-913">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="4e51e-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4e51e-914">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4e51e-915">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-915">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4e51e-916">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-916">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4e51e-917">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="4e51e-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4e51e-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="4e51e-919">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-920">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-921">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="4e51e-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4e51e-922">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4e51e-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4e51e-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-926">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-926">Parameters:</span></span>

|<span data-ttu-id="4e51e-927">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-927">Name</span></span>|<span data-ttu-id="4e51e-928">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-928">Type</span></span>|<span data-ttu-id="4e51e-929">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-929">Attributes</span></span>|<span data-ttu-id="4e51e-930">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4e51e-931">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4e51e-931">String &#124; Object</span></span>||<span data-ttu-id="4e51e-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4e51e-934">**OU**</span><span class="sxs-lookup"><span data-stu-id="4e51e-934">**OR**</span></span><br/><span data-ttu-id="4e51e-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4e51e-937">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-937">String</span></span>|<span data-ttu-id="4e51e-938">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-938">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4e51e-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4e51e-942">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-942">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-943">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4e51e-944">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-944">String</span></span>||<span data-ttu-id="4e51e-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4e51e-947">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-947">String</span></span>||<span data-ttu-id="4e51e-948">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="4e51e-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4e51e-949">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-949">String</span></span>||<span data-ttu-id="4e51e-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4e51e-952">Boolean</span><span class="sxs-lookup"><span data-stu-id="4e51e-952">Boolean</span></span>||<span data-ttu-id="4e51e-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4e51e-955">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-955">String</span></span>||<span data-ttu-id="4e51e-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4e51e-959">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-959">function</span></span>|<span data-ttu-id="4e51e-960">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-960">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-961">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-962">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-962">Requirements</span></span>

|<span data-ttu-id="4e51e-963">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-963">Requirement</span></span>|<span data-ttu-id="4e51e-964">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-965">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-966">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-966">1.0</span></span>|
|[<span data-ttu-id="4e51e-967">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-968">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-969">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-970">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e51e-971">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4e51e-971">Examples</span></span>

<span data-ttu-id="4e51e-972">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4e51e-973">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="4e51e-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4e51e-974">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4e51e-975">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-975">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4e51e-976">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-976">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4e51e-977">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4e51e-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4e51e-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4e51e-979">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-980">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-981">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-981">Requirements</span></span>

|<span data-ttu-id="4e51e-982">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-982">Requirement</span></span>|<span data-ttu-id="4e51e-983">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-984">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-985">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-985">1.0</span></span>|
|[<span data-ttu-id="4e51e-986">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-987">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-988">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-989">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-990">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-990">Returns:</span></span>

<span data-ttu-id="4e51e-991">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4e51e-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4e51e-992">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-992">Example</span></span>

<span data-ttu-id="4e51e-993">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4e51e-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4e51e-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4e51e-995">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-996">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-997">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-997">Parameters:</span></span>

|<span data-ttu-id="4e51e-998">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-998">Name</span></span>|<span data-ttu-id="4e51e-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-999">Type</span></span>|<span data-ttu-id="4e51e-1000">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4e51e-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4e51e-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="4e51e-1002">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1003">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1003">Requirements</span></span>

|<span data-ttu-id="4e51e-1004">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1004">Requirement</span></span>|<span data-ttu-id="4e51e-1005">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1006">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-1007">1.0</span></span>|
|[<span data-ttu-id="4e51e-1008">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1009">Restrito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1009">Restricted</span></span>|
|[<span data-ttu-id="4e51e-1010">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1011">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1012">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1012">Returns:</span></span>

<span data-ttu-id="4e51e-1013">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4e51e-1014">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4e51e-1015">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4e51e-1016">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4e51e-1017">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4e51e-1017">Value of `entityType`</span></span>|<span data-ttu-id="4e51e-1018">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="4e51e-1018">Type of objects in returned array</span></span>|<span data-ttu-id="4e51e-1019">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="4e51e-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4e51e-1020">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1020">String</span></span>|<span data-ttu-id="4e51e-1021">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4e51e-1022">Contato</span><span class="sxs-lookup"><span data-stu-id="4e51e-1022">Contact</span></span>|<span data-ttu-id="4e51e-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4e51e-1024">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1024">String</span></span>|<span data-ttu-id="4e51e-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4e51e-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4e51e-1026">MeetingSuggestion</span></span>|<span data-ttu-id="4e51e-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4e51e-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4e51e-1028">PhoneNumber</span></span>|<span data-ttu-id="4e51e-1029">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4e51e-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4e51e-1030">TaskSuggestion</span></span>|<span data-ttu-id="4e51e-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4e51e-1032">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1032">String</span></span>|<span data-ttu-id="4e51e-1033">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4e51e-1033">**Restricted**</span></span>|

<span data-ttu-id="4e51e-1034">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4e51e-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4e51e-1035">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1035">Example</span></span>

<span data-ttu-id="4e51e-1036">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4e51e-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4e51e-1038">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1039">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-1040">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1041">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1041">Parameters:</span></span>

|<span data-ttu-id="4e51e-1042">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1042">Name</span></span>|<span data-ttu-id="4e51e-1043">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1043">Type</span></span>|<span data-ttu-id="4e51e-1044">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4e51e-1045">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1045">String</span></span>|<span data-ttu-id="4e51e-1046">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1047">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1047">Requirements</span></span>

|<span data-ttu-id="4e51e-1048">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1048">Requirement</span></span>|<span data-ttu-id="4e51e-1049">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1050">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-1051">1.0</span></span>|
|[<span data-ttu-id="4e51e-1052">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1053">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1054">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1055">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1056">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1056">Returns:</span></span>

<span data-ttu-id="4e51e-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4e51e-1059">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4e51e-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="4e51e-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e51e-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="4e51e-1061">Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1062">Este método só é suportado pelo Outlook 2016 para Windows (Click-to-Run versões maiores 16.0.8413.1000) e do Outlook na web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1063">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1063">Parameters:</span></span>
|<span data-ttu-id="4e51e-1064">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1064">Name</span></span>|<span data-ttu-id="4e51e-1065">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1065">Type</span></span>|<span data-ttu-id="4e51e-1066">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1066">Attributes</span></span>|<span data-ttu-id="4e51e-1067">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4e51e-1068">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1068">Object</span></span>|<span data-ttu-id="4e51e-1069">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1070">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-1071">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1071">Object</span></span>|<span data-ttu-id="4e51e-1072">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1073">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4e51e-1074">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1074">function</span></span>|<span data-ttu-id="4e51e-1075">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1076">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e51e-1077">Em caso de sucesso, os dados de inicialização são fornecidos no `asyncResult.value` a propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="4e51e-1078">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1079">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1079">Requirements</span></span>

|<span data-ttu-id="4e51e-1080">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1080">Requirement</span></span>|<span data-ttu-id="4e51e-1081">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1082">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1083">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-1083">Preview</span></span>|
|[<span data-ttu-id="4e51e-1084">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1085">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1086">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1087">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-1088">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1088">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="4e51e-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4e51e-1090">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1091">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4e51e-1095">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4e51e-1096">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4e51e-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-1100">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1100">Requirements</span></span>

|<span data-ttu-id="4e51e-1101">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1101">Requirement</span></span>|<span data-ttu-id="4e51e-1102">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1103">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-1104">1.0</span></span>|
|[<span data-ttu-id="4e51e-1105">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1106">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1107">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1108">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1109">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1109">Returns:</span></span>

<span data-ttu-id="4e51e-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4e51e-1112">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4e51e-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4e51e-1113">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4e51e-1114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1114">Example</span></span>

<span data-ttu-id="4e51e-1115">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4e51e-1116">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4e51e-1117">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1118">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-1119">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4e51e-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1122">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1122">Parameters:</span></span>

|<span data-ttu-id="4e51e-1123">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1123">Name</span></span>|<span data-ttu-id="4e51e-1124">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1124">Type</span></span>|<span data-ttu-id="4e51e-1125">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4e51e-1126">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1126">String</span></span>|<span data-ttu-id="4e51e-1127">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1128">Requirements</span></span>

|<span data-ttu-id="4e51e-1129">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1129">Requirement</span></span>|<span data-ttu-id="4e51e-1130">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-1132">1.0</span></span>|
|[<span data-ttu-id="4e51e-1133">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1134">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1135">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1136">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1137">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1137">Returns:</span></span>

<span data-ttu-id="4e51e-1138">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4e51e-1139">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="4e51e-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4e51e-1140">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="4e51e-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4e51e-1141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4e51e-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4e51e-1143">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4e51e-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1146">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1146">Parameters:</span></span>

|<span data-ttu-id="4e51e-1147">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1147">Name</span></span>|<span data-ttu-id="4e51e-1148">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1148">Type</span></span>|<span data-ttu-id="4e51e-1149">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1149">Attributes</span></span>|<span data-ttu-id="4e51e-1150">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4e51e-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4e51e-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4e51e-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4e51e-1155">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1155">Object</span></span>|<span data-ttu-id="4e51e-1156">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1157">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-1158">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1158">Object</span></span>|<span data-ttu-id="4e51e-1159">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1160">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4e51e-1161">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1161">function</span></span>||<span data-ttu-id="4e51e-1162">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e51e-1163">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4e51e-1164">Para acessar a propriedade source que a seleção é proveniente, chamar `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1165">Requirements</span></span>

|<span data-ttu-id="4e51e-1166">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1166">Requirement</span></span>|<span data-ttu-id="4e51e-1167">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="4e51e-1169">1.2</span></span>|
|[<span data-ttu-id="4e51e-1170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-1172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1173">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1174">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1174">Returns:</span></span>

<span data-ttu-id="4e51e-1175">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4e51e-1176">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4e51e-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4e51e-1177">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4e51e-1178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1178">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4e51e-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4e51e-p168">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4e51e-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1182">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-1183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1183">Requirements</span></span>

|<span data-ttu-id="4e51e-1184">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1184">Requirement</span></span>|<span data-ttu-id="4e51e-1185">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="4e51e-1187">1.6</span></span>|
|[<span data-ttu-id="4e51e-1188">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1189">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1191">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1192">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1192">Returns:</span></span>

<span data-ttu-id="4e51e-1193">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4e51e-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4e51e-1194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1194">Example</span></span>

<span data-ttu-id="4e51e-1195">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4e51e-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4e51e-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4e51e-p169">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4e51e-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1199">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4e51e-p170">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4e51e-1203">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4e51e-1204">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4e51e-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e51e-1208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1208">Requirements</span></span>

|<span data-ttu-id="4e51e-1209">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1209">Requirement</span></span>|<span data-ttu-id="4e51e-1210">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="4e51e-1212">1.6</span></span>|
|[<span data-ttu-id="4e51e-1213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1214">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1216">Leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e51e-1217">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1217">Returns:</span></span>

<span data-ttu-id="4e51e-p172">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4e51e-1220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1220">Example</span></span>

<span data-ttu-id="4e51e-1221">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4e51e-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4e51e-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4e51e-1223">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4e51e-p173">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1227">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1227">Parameters:</span></span>

|<span data-ttu-id="4e51e-1228">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1228">Name</span></span>|<span data-ttu-id="4e51e-1229">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1229">Type</span></span>|<span data-ttu-id="4e51e-1230">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1230">Attributes</span></span>|<span data-ttu-id="4e51e-1231">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4e51e-1232">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1232">function</span></span>||<span data-ttu-id="4e51e-1233">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e51e-1234">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4e51e-1235">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4e51e-1236">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1236">Object</span></span>|<span data-ttu-id="4e51e-1237">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1238">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4e51e-1239">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1240">Requirements</span></span>

|<span data-ttu-id="4e51e-1241">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1241">Requirement</span></span>|<span data-ttu-id="4e51e-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="4e51e-1244">1.0</span></span>|
|[<span data-ttu-id="4e51e-1245">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1246">ReadItem</span></span>|
|[<span data-ttu-id="4e51e-1247">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1248">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-1249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1249">Example</span></span>

<span data-ttu-id="4e51e-p176">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4e51e-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e51e-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4e51e-1254">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4e51e-p177">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1259">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1259">Parameters:</span></span>

|<span data-ttu-id="4e51e-1260">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1260">Name</span></span>|<span data-ttu-id="4e51e-1261">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1261">Type</span></span>|<span data-ttu-id="4e51e-1262">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1262">Attributes</span></span>|<span data-ttu-id="4e51e-1263">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4e51e-1264">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4e51e-1264">String</span></span>||<span data-ttu-id="4e51e-p178">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="4e51e-1267">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1267">Object</span></span>|<span data-ttu-id="4e51e-1268">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1269">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-1270">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1270">Object</span></span>|<span data-ttu-id="4e51e-1271">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1272">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4e51e-1273">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1273">function</span></span>|<span data-ttu-id="4e51e-1274">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1275">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e51e-1276">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e51e-1277">Erros</span><span class="sxs-lookup"><span data-stu-id="4e51e-1277">Errors</span></span>

|<span data-ttu-id="4e51e-1278">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4e51e-1278">Error code</span></span>|<span data-ttu-id="4e51e-1279">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4e51e-1280">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1281">Requirements</span></span>

|<span data-ttu-id="4e51e-1282">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1282">Requirement</span></span>|<span data-ttu-id="4e51e-1283">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="4e51e-1285">1.1</span></span>|
|[<span data-ttu-id="4e51e-1286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-1288">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1289">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-1290">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1290">Example</span></span>

<span data-ttu-id="4e51e-1291">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1291">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4e51e-1292">removeHandlerAsync (eventType, manipulador, [Opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="4e51e-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4e51e-1293">Remove um manipulador de eventos para um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="4e51e-1294">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="4e51e-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1295">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1295">Parameters:</span></span>

| <span data-ttu-id="4e51e-1296">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1296">Name</span></span> | <span data-ttu-id="4e51e-1297">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1297">Type</span></span> | <span data-ttu-id="4e51e-1298">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1298">Attributes</span></span> | <span data-ttu-id="4e51e-1299">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4e51e-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4e51e-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4e51e-1301">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4e51e-1302">Função</span><span class="sxs-lookup"><span data-stu-id="4e51e-1302">Function</span></span> || <span data-ttu-id="4e51e-p179">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4e51e-1306">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1306">Object</span></span> | <span data-ttu-id="4e51e-1307">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="4e51e-1308">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4e51e-1309">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1309">Object</span></span> | <span data-ttu-id="4e51e-1310">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="4e51e-1311">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4e51e-1312">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1312">function</span></span>| <span data-ttu-id="4e51e-1313">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1314">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1315">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1315">Requirements</span></span>

|<span data-ttu-id="4e51e-1316">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1316">Requirement</span></span>| <span data-ttu-id="4e51e-1317">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1318">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e51e-1319">Visualização</span><span class="sxs-lookup"><span data-stu-id="4e51e-1319">Preview</span></span> |
|[<span data-ttu-id="4e51e-1320">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e51e-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1321">ReadItem</span></span> |
|[<span data-ttu-id="4e51e-1322">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e51e-1323">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="4e51e-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="4e51e-1324">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="4e51e-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4e51e-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="4e51e-1326">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="4e51e-p180">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1330">Se seu suplemento chama `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, pode levar algum tempo antes do item realmente é sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4e51e-1331">Até que o item é sincronizado, usando o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4e51e-p182">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4e51e-1335">Os seguintes clientes tenham comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4e51e-1336">Mac Outlook não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="4e51e-1337">Chamar `saveAsync` em uma reunião no Outlook Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="4e51e-1338">Outlook na web sempre envia um convite ou atualizar quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1339">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1339">Parameters:</span></span>

|<span data-ttu-id="4e51e-1340">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1340">Name</span></span>|<span data-ttu-id="4e51e-1341">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1341">Type</span></span>|<span data-ttu-id="4e51e-1342">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1342">Attributes</span></span>|<span data-ttu-id="4e51e-1343">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4e51e-1344">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1344">Object</span></span>|<span data-ttu-id="4e51e-1345">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1346">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-1347">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1347">Object</span></span>|<span data-ttu-id="4e51e-1348">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1349">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4e51e-1350">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1350">function</span></span>||<span data-ttu-id="4e51e-1351">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e51e-1352">Em caso de sucesso, o identificador do item é fornecido no `asyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1353">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1353">Requirements</span></span>

|<span data-ttu-id="4e51e-1354">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1354">Requirement</span></span>|<span data-ttu-id="4e51e-1355">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1356">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="4e51e-1357">1.3</span></span>|
|[<span data-ttu-id="4e51e-1358">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-1360">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1361">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e51e-1362">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="4e51e-p184">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4e51e-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4e51e-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4e51e-1366">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4e51e-p185">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e51e-1370">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="4e51e-1370">Parameters:</span></span>

|<span data-ttu-id="4e51e-1371">Nome</span><span class="sxs-lookup"><span data-stu-id="4e51e-1371">Name</span></span>|<span data-ttu-id="4e51e-1372">Tipo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1372">Type</span></span>|<span data-ttu-id="4e51e-1373">Atributos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1373">Attributes</span></span>|<span data-ttu-id="4e51e-1374">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4e51e-1375">String</span><span class="sxs-lookup"><span data-stu-id="4e51e-1375">String</span></span>||<span data-ttu-id="4e51e-p186">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4e51e-1379">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1379">Object</span></span>|<span data-ttu-id="4e51e-1380">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1381">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4e51e-1382">Objeto</span><span class="sxs-lookup"><span data-stu-id="4e51e-1382">Object</span></span>|<span data-ttu-id="4e51e-1383">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-1384">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4e51e-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4e51e-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4e51e-1386">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e51e-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="4e51e-p187">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4e51e-p188">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="4e51e-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4e51e-1391">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="4e51e-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4e51e-1392">function</span><span class="sxs-lookup"><span data-stu-id="4e51e-1392">function</span></span>||<span data-ttu-id="4e51e-1393">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4e51e-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e51e-1394">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4e51e-1394">Requirements</span></span>

|<span data-ttu-id="4e51e-1395">Requisito</span><span class="sxs-lookup"><span data-stu-id="4e51e-1395">Requirement</span></span>|<span data-ttu-id="4e51e-1396">Valor</span><span class="sxs-lookup"><span data-stu-id="4e51e-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e51e-1397">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4e51e-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4e51e-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="4e51e-1398">1.2</span></span>|
|[<span data-ttu-id="4e51e-1399">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4e51e-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e51e-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e51e-1401">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4e51e-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="4e51e-1402">Composição</span><span class="sxs-lookup"><span data-stu-id="4e51e-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e51e-1403">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4e51e-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```