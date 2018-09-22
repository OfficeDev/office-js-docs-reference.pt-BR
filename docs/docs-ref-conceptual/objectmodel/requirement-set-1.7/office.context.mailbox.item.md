
# <a name="item"></a><span data-ttu-id="00e05-101">item</span><span class="sxs-lookup"><span data-stu-id="00e05-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="00e05-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="00e05-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="00e05-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="00e05-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-105">Requirements</span></span>

|<span data-ttu-id="00e05-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-106">Requirement</span></span>|<span data-ttu-id="00e05-107">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-109">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-109">1.0</span></span>|
|[<span data-ttu-id="00e05-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="00e05-111">Restricted</span></span>|
|[<span data-ttu-id="00e05-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="00e05-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="00e05-114">Members and methods</span></span>

| <span data-ttu-id="00e05-115">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-115">Member</span></span> | <span data-ttu-id="00e05-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="00e05-117">attachments</span><span class="sxs-lookup"><span data-stu-id="00e05-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="00e05-118">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-118">Member</span></span> |
| [<span data-ttu-id="00e05-119">bcc</span><span class="sxs-lookup"><span data-stu-id="00e05-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="00e05-120">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-120">Member</span></span> |
| [<span data-ttu-id="00e05-121">body</span><span class="sxs-lookup"><span data-stu-id="00e05-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="00e05-122">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-122">Member</span></span> |
| [<span data-ttu-id="00e05-123">cc</span><span class="sxs-lookup"><span data-stu-id="00e05-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="00e05-124">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-124">Member</span></span> |
| [<span data-ttu-id="00e05-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="00e05-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="00e05-126">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-126">Member</span></span> |
| [<span data-ttu-id="00e05-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="00e05-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="00e05-128">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-128">Member</span></span> |
| [<span data-ttu-id="00e05-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="00e05-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="00e05-130">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-130">Member</span></span> |
| [<span data-ttu-id="00e05-131">end</span><span class="sxs-lookup"><span data-stu-id="00e05-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="00e05-132">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-132">Member</span></span> |
| [<span data-ttu-id="00e05-133">from</span><span class="sxs-lookup"><span data-stu-id="00e05-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="00e05-134">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-134">Member</span></span> |
| [<span data-ttu-id="00e05-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="00e05-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="00e05-136">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-136">Member</span></span> |
| [<span data-ttu-id="00e05-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="00e05-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="00e05-138">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-138">Member</span></span> |
| [<span data-ttu-id="00e05-139">itemId</span><span class="sxs-lookup"><span data-stu-id="00e05-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="00e05-140">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-140">Member</span></span> |
| [<span data-ttu-id="00e05-141">itemType</span><span class="sxs-lookup"><span data-stu-id="00e05-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="00e05-142">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-142">Member</span></span> |
| [<span data-ttu-id="00e05-143">location</span><span class="sxs-lookup"><span data-stu-id="00e05-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="00e05-144">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-144">Member</span></span> |
| [<span data-ttu-id="00e05-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="00e05-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="00e05-146">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-146">Member</span></span> |
| [<span data-ttu-id="00e05-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="00e05-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="00e05-148">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-148">Member</span></span> |
| [<span data-ttu-id="00e05-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="00e05-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="00e05-150">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-150">Member</span></span> |
| [<span data-ttu-id="00e05-151">organizer</span><span class="sxs-lookup"><span data-stu-id="00e05-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="00e05-152">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-152">Member</span></span> |
| [<span data-ttu-id="00e05-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="00e05-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="00e05-154">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-154">Member</span></span> |
| [<span data-ttu-id="00e05-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="00e05-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="00e05-156">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-156">Member</span></span> |
| [<span data-ttu-id="00e05-157">sender</span><span class="sxs-lookup"><span data-stu-id="00e05-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="00e05-158">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-158">Member</span></span> |
| [<span data-ttu-id="00e05-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="00e05-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="00e05-160">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-160">Member</span></span> |
| [<span data-ttu-id="00e05-161">start</span><span class="sxs-lookup"><span data-stu-id="00e05-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="00e05-162">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-162">Member</span></span> |
| [<span data-ttu-id="00e05-163">subject</span><span class="sxs-lookup"><span data-stu-id="00e05-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="00e05-164">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-164">Member</span></span> |
| [<span data-ttu-id="00e05-165">to</span><span class="sxs-lookup"><span data-stu-id="00e05-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="00e05-166">Membro</span><span class="sxs-lookup"><span data-stu-id="00e05-166">Member</span></span> |
| [<span data-ttu-id="00e05-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="00e05-168">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-168">Method</span></span> |
| [<span data-ttu-id="00e05-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="00e05-170">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-170">Method</span></span> |
| [<span data-ttu-id="00e05-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="00e05-172">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-172">Method</span></span> |
| [<span data-ttu-id="00e05-173">close</span><span class="sxs-lookup"><span data-stu-id="00e05-173">close</span></span>](#close) | <span data-ttu-id="00e05-174">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-174">Method</span></span> |
| [<span data-ttu-id="00e05-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="00e05-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="00e05-176">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-176">Method</span></span> |
| [<span data-ttu-id="00e05-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="00e05-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="00e05-178">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-178">Method</span></span> |
| [<span data-ttu-id="00e05-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="00e05-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="00e05-180">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-180">Method</span></span> |
| [<span data-ttu-id="00e05-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="00e05-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="00e05-182">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-182">Method</span></span> |
| [<span data-ttu-id="00e05-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="00e05-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="00e05-184">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-184">Method</span></span> |
| [<span data-ttu-id="00e05-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="00e05-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="00e05-186">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-186">Method</span></span> |
| [<span data-ttu-id="00e05-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="00e05-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="00e05-188">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-188">Method</span></span> |
| [<span data-ttu-id="00e05-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="00e05-190">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-190">Method</span></span> |
| [<span data-ttu-id="00e05-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="00e05-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="00e05-192">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-192">Method</span></span> |
| [<span data-ttu-id="00e05-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="00e05-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="00e05-194">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-194">Method</span></span> |
| [<span data-ttu-id="00e05-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="00e05-196">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-196">Method</span></span> |
| [<span data-ttu-id="00e05-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="00e05-198">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-198">Method</span></span> |
| [<span data-ttu-id="00e05-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="00e05-200">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-200">Method</span></span> |
| [<span data-ttu-id="00e05-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="00e05-202">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-202">Method</span></span> |
| [<span data-ttu-id="00e05-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00e05-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="00e05-204">Método</span><span class="sxs-lookup"><span data-stu-id="00e05-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="00e05-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-205">Example</span></span>

<span data-ttu-id="00e05-206">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="00e05-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="00e05-207">Membros</span><span class="sxs-lookup"><span data-stu-id="00e05-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="00e05-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00e05-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="00e05-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-211">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="00e05-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="00e05-212">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="00e05-212">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-213">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-213">Type:</span></span>

*   <span data-ttu-id="00e05-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00e05-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-215">Requirements</span></span>

|<span data-ttu-id="00e05-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-216">Requirement</span></span>|<span data-ttu-id="00e05-217">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-219">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-219">1.0</span></span>|
|[<span data-ttu-id="00e05-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-221">ReadItem</span></span>|
|[<span data-ttu-id="00e05-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-223">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-224">Example</span></span>

<span data-ttu-id="00e05-225">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="00e05-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="00e05-227">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="00e05-228">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="00e05-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-229">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-229">Type:</span></span>

*   [<span data-ttu-id="00e05-230">Destinatários</span><span class="sxs-lookup"><span data-stu-id="00e05-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="00e05-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-231">Requirements</span></span>

|<span data-ttu-id="00e05-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-232">Requirement</span></span>|<span data-ttu-id="00e05-233">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-235">1.1</span><span class="sxs-lookup"><span data-stu-id="00e05-235">1.1</span></span>|
|[<span data-ttu-id="00e05-236">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-237">ReadItem</span></span>|
|[<span data-ttu-id="00e05-238">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-239">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-240">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="00e05-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="00e05-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="00e05-242">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="00e05-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-243">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-243">Type:</span></span>

*   [<span data-ttu-id="00e05-244">Corpo</span><span class="sxs-lookup"><span data-stu-id="00e05-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="00e05-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-245">Requirements</span></span>

|<span data-ttu-id="00e05-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-246">Requirement</span></span>|<span data-ttu-id="00e05-247">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-249">1.1</span><span class="sxs-lookup"><span data-stu-id="00e05-249">1.1</span></span>|
|[<span data-ttu-id="00e05-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-251">ReadItem</span></span>|
|[<span data-ttu-id="00e05-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-253">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="00e05-254">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="00e05-255">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="00e05-256">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-257">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-257">Read mode</span></span>

<span data-ttu-id="00e05-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="00e05-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-260">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-260">Compose mode</span></span>

<span data-ttu-id="00e05-261">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-261">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-262">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-262">Type:</span></span>

*   <span data-ttu-id="00e05-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-264">Requirements</span></span>

|<span data-ttu-id="00e05-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-265">Requirement</span></span>|<span data-ttu-id="00e05-266">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-268">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-268">1.0</span></span>|
|[<span data-ttu-id="00e05-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-270">ReadItem</span></span>|
|[<span data-ttu-id="00e05-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-272">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="00e05-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="00e05-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="00e05-275">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="00e05-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="00e05-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="00e05-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="00e05-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="00e05-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-280">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-280">Type:</span></span>

*   <span data-ttu-id="00e05-281">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-282">Requirements</span></span>

|<span data-ttu-id="00e05-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-283">Requirement</span></span>|<span data-ttu-id="00e05-284">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-286">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-286">1.0</span></span>|
|[<span data-ttu-id="00e05-287">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-288">ReadItem</span></span>|
|[<span data-ttu-id="00e05-289">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-290">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="00e05-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="00e05-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="00e05-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-294">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-294">Type:</span></span>

*   <span data-ttu-id="00e05-295">Data</span><span class="sxs-lookup"><span data-stu-id="00e05-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-296">Requirements</span></span>

|<span data-ttu-id="00e05-297">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-297">Requirement</span></span>|<span data-ttu-id="00e05-298">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-299">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-300">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-300">1.0</span></span>|
|[<span data-ttu-id="00e05-301">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-302">ReadItem</span></span>|
|[<span data-ttu-id="00e05-303">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-304">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-305">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="00e05-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="00e05-306">dateTimeModified :Date</span></span>

<span data-ttu-id="00e05-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-309">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-309">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-310">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-310">Type:</span></span>

*   <span data-ttu-id="00e05-311">Data</span><span class="sxs-lookup"><span data-stu-id="00e05-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-312">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-312">Requirements</span></span>

|<span data-ttu-id="00e05-313">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-313">Requirement</span></span>|<span data-ttu-id="00e05-314">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-315">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-316">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-316">1.0</span></span>|
|[<span data-ttu-id="00e05-317">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-318">ReadItem</span></span>|
|[<span data-ttu-id="00e05-319">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-320">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-321">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="00e05-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="00e05-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="00e05-323">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="00e05-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="00e05-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="00e05-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-326">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-326">Read mode</span></span>

<span data-ttu-id="00e05-327">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="00e05-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-328">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-328">Compose mode</span></span>

<span data-ttu-id="00e05-329">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="00e05-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="00e05-330">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="00e05-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-331">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-331">Type:</span></span>

*   <span data-ttu-id="00e05-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="00e05-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-333">Requirements</span></span>

|<span data-ttu-id="00e05-334">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-334">Requirement</span></span>|<span data-ttu-id="00e05-335">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-336">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-337">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-337">1.0</span></span>|
|[<span data-ttu-id="00e05-338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-339">ReadItem</span></span>|
|[<span data-ttu-id="00e05-340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-341">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-342">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-342">Example</span></span>

<span data-ttu-id="00e05-343">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="00e05-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="00e05-344">de:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="00e05-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="00e05-345">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="00e05-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="00e05-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-348">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00e05-348">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-349">Read mode</span></span>

<span data-ttu-id="00e05-350">O `from` propriedade retornará uma `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="00e05-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="00e05-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-351">Compose mode</span></span>

<span data-ttu-id="00e05-352">O `from` propriedade retornará uma `From` objeto que fornece um método para obter o de valor.</span><span class="sxs-lookup"><span data-stu-id="00e05-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00e05-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-353">Type:</span></span>

*   <span data-ttu-id="00e05-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="00e05-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-355">Requirements</span></span>

|<span data-ttu-id="00e05-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="00e05-357">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-358">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-358">1.0</span></span>|<span data-ttu-id="00e05-359">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-359">1.7</span></span>|
|[<span data-ttu-id="00e05-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-361">ReadItem</span></span>|<span data-ttu-id="00e05-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-363">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-364">Read</span><span class="sxs-lookup"><span data-stu-id="00e05-364">Read</span></span>|<span data-ttu-id="00e05-365">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="00e05-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="00e05-366">internetMessageId :String</span></span>

<span data-ttu-id="00e05-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-369">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-369">Type:</span></span>

*   <span data-ttu-id="00e05-370">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-371">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-371">Requirements</span></span>

|<span data-ttu-id="00e05-372">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-372">Requirement</span></span>|<span data-ttu-id="00e05-373">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-374">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-375">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-375">1.0</span></span>|
|[<span data-ttu-id="00e05-376">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-377">ReadItem</span></span>|
|[<span data-ttu-id="00e05-378">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-379">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-380">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="00e05-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="00e05-381">itemClass :String</span></span>

<span data-ttu-id="00e05-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="00e05-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="00e05-386">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-386">Type</span></span>|<span data-ttu-id="00e05-387">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-387">Description</span></span>|<span data-ttu-id="00e05-388">classe de item</span><span class="sxs-lookup"><span data-stu-id="00e05-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="00e05-389">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="00e05-389">Appointment items</span></span>|<span data-ttu-id="00e05-390">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="00e05-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="00e05-391">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="00e05-391">Message items</span></span>|<span data-ttu-id="00e05-392">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="00e05-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="00e05-393">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="00e05-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-394">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-394">Type:</span></span>

*   <span data-ttu-id="00e05-395">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-396">Requirements</span></span>

|<span data-ttu-id="00e05-397">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-397">Requirement</span></span>|<span data-ttu-id="00e05-398">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-399">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-400">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-400">1.0</span></span>|
|[<span data-ttu-id="00e05-401">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-402">ReadItem</span></span>|
|[<span data-ttu-id="00e05-403">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-404">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-405">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="00e05-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="00e05-406">(nullable) itemId :String</span></span>

<span data-ttu-id="00e05-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-409">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="00e05-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="00e05-410">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="00e05-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="00e05-411">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="00e05-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="00e05-412">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="00e05-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="00e05-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-415">Type:</span></span>

*   <span data-ttu-id="00e05-416">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-417">Requirements</span></span>

|<span data-ttu-id="00e05-418">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-418">Requirement</span></span>|<span data-ttu-id="00e05-419">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-420">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-421">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-421">1.0</span></span>|
|[<span data-ttu-id="00e05-422">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-423">ReadItem</span></span>|
|[<span data-ttu-id="00e05-424">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-425">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-426">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-426">Example</span></span>

<span data-ttu-id="00e05-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="00e05-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="00e05-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="00e05-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="00e05-430">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="00e05-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="00e05-431">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-432">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-432">Type:</span></span>

*   [<span data-ttu-id="00e05-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="00e05-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="00e05-434">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-434">Requirements</span></span>

|<span data-ttu-id="00e05-435">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-435">Requirement</span></span>|<span data-ttu-id="00e05-436">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-437">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-438">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-438">1.0</span></span>|
|[<span data-ttu-id="00e05-439">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-440">ReadItem</span></span>|
|[<span data-ttu-id="00e05-441">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-442">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-443">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="00e05-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="00e05-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="00e05-445">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-446">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-446">Read mode</span></span>

<span data-ttu-id="00e05-447">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-448">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-448">Compose mode</span></span>

<span data-ttu-id="00e05-449">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-450">Type:</span></span>

*   <span data-ttu-id="00e05-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="00e05-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-452">Requirements</span></span>

|<span data-ttu-id="00e05-453">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-453">Requirement</span></span>|<span data-ttu-id="00e05-454">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-455">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-456">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-456">1.0</span></span>|
|[<span data-ttu-id="00e05-457">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-458">ReadItem</span></span>|
|[<span data-ttu-id="00e05-459">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-460">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-461">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="00e05-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="00e05-462">normalizedSubject :String</span></span>

<span data-ttu-id="00e05-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="00e05-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="00e05-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-467">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-467">Type:</span></span>

*   <span data-ttu-id="00e05-468">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-469">Requirements</span></span>

|<span data-ttu-id="00e05-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-470">Requirement</span></span>|<span data-ttu-id="00e05-471">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-473">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-473">1.0</span></span>|
|[<span data-ttu-id="00e05-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-475">ReadItem</span></span>|
|[<span data-ttu-id="00e05-476">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-477">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="00e05-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="00e05-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="00e05-480">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="00e05-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-481">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-481">Type:</span></span>

*   [<span data-ttu-id="00e05-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="00e05-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="00e05-483">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-483">Requirements</span></span>

|<span data-ttu-id="00e05-484">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-484">Requirement</span></span>|<span data-ttu-id="00e05-485">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-486">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-487">1.3</span><span class="sxs-lookup"><span data-stu-id="00e05-487">1.3</span></span>|
|[<span data-ttu-id="00e05-488">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-489">ReadItem</span></span>|
|[<span data-ttu-id="00e05-490">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-491">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="00e05-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="00e05-493">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="00e05-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="00e05-494">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-495">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-495">Read mode</span></span>

<span data-ttu-id="00e05-496">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-497">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-497">Compose mode</span></span>

<span data-ttu-id="00e05-498">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-499">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-499">Type:</span></span>

*   <span data-ttu-id="00e05-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-501">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-501">Requirements</span></span>

|<span data-ttu-id="00e05-502">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-502">Requirement</span></span>|<span data-ttu-id="00e05-503">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-504">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-505">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-505">1.0</span></span>|
|[<span data-ttu-id="00e05-506">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-507">ReadItem</span></span>|
|[<span data-ttu-id="00e05-508">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-509">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-510">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="00e05-511">Organizador:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizador](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="00e05-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="00e05-512">Obtém o endereço de email do organizador da reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="00e05-512">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-513">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-513">Read mode</span></span>

<span data-ttu-id="00e05-514">O `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-515">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-515">Compose mode</span></span>

<span data-ttu-id="00e05-516">O `organizer` propriedade retorna um objeto do [Organizador](/javascript/api/outlook_1_7/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="00e05-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-517">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-517">Type:</span></span>

*   <span data-ttu-id="00e05-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizador](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="00e05-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-519">Requirements</span></span>

|<span data-ttu-id="00e05-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="00e05-521">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-522">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-522">1.0</span></span>|<span data-ttu-id="00e05-523">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-523">1.7</span></span>|
|[<span data-ttu-id="00e05-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-525">ReadItem</span></span>|<span data-ttu-id="00e05-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-527">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-528">Read</span><span class="sxs-lookup"><span data-stu-id="00e05-528">Read</span></span>|<span data-ttu-id="00e05-529">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-530">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="00e05-531">Recorrência (anulável):[Recorrência](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="00e05-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="00e05-532">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-532">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="00e05-533">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="00e05-534">Leitura e redação modos para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="00e05-535">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="00e05-536">O `recurrence` propriedade retorna um objeto de [Recorrência](/javascript/api/outlook_1_7/office.recurrence) para solicitações de reuniões ou compromissos recorrentes, se um item é uma série ou uma instância de uma série.</span><span class="sxs-lookup"><span data-stu-id="00e05-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="00e05-537">`null`é retornada para um único compromissos e solicitações de reunião de compromissos de única.</span><span class="sxs-lookup"><span data-stu-id="00e05-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="00e05-538">`undefined`é retornado para mensagens que não fazem solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="00e05-539">Observação: Solicitações de reunião tem um `itemClass` valor IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="00e05-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="00e05-540">Observação: Se o objeto de recorrência for `null`, isto indica que o objeto é um compromisso único ou uma solicitação de reunião de um único compromisso e não fizer parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="00e05-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-541">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-541">Type:</span></span>

* [<span data-ttu-id="00e05-542">Recorrência</span><span class="sxs-lookup"><span data-stu-id="00e05-542">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="00e05-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-543">Requirement</span></span>|<span data-ttu-id="00e05-544">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-546">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-546">1.7</span></span>|
|[<span data-ttu-id="00e05-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-548">ReadItem</span></span>|
|[<span data-ttu-id="00e05-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-550">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="00e05-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="00e05-552">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="00e05-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="00e05-553">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-554">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-554">Read mode</span></span>

<span data-ttu-id="00e05-555">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-556">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-556">Compose mode</span></span>

<span data-ttu-id="00e05-557">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-558">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-558">Type:</span></span>

*   <span data-ttu-id="00e05-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-560">Requirements</span></span>

|<span data-ttu-id="00e05-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-561">Requirement</span></span>|<span data-ttu-id="00e05-562">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-564">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-564">1.0</span></span>|
|[<span data-ttu-id="00e05-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-566">ReadItem</span></span>|
|[<span data-ttu-id="00e05-567">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-568">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-569">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="00e05-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00e05-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="00e05-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="00e05-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="00e05-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="00e05-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-575">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `sender` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00e05-575">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-576">Type:</span></span>

*   [<span data-ttu-id="00e05-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00e05-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00e05-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-578">Requirements</span></span>

|<span data-ttu-id="00e05-579">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-579">Requirement</span></span>|<span data-ttu-id="00e05-580">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-581">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-582">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-582">1.0</span></span>|
|[<span data-ttu-id="00e05-583">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-584">ReadItem</span></span>|
|[<span data-ttu-id="00e05-585">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-586">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-587">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="00e05-588">seriesId (anulável): cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="00e05-589">Obtém a identificação da série que uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="00e05-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="00e05-590">No OWA e no Outlook, o `seriesId` retornará a identificação de serviços Web do Exchange (EWS) do item pai (série) que este item pertence.</span><span class="sxs-lookup"><span data-stu-id="00e05-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="00e05-591">No entanto, em iOS e Android, o `seriesId` retornará a identificação do restante do item pai.</span><span class="sxs-lookup"><span data-stu-id="00e05-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-592">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="00e05-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="00e05-593">O `seriesId` propriedade não é idêntica às IDs do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="00e05-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="00e05-594">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="00e05-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="00e05-595">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="00e05-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="00e05-596">O `seriesId` propriedade retornará `null` para itens que não têm itens pai como único compromissos, itens de série, ou solicitações de reunião e retorna `undefined` para todos os itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="00e05-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-597">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-597">Type:</span></span>

* <span data-ttu-id="00e05-598">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-599">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-599">Requirements</span></span>

|<span data-ttu-id="00e05-600">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-600">Requirement</span></span>|<span data-ttu-id="00e05-601">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-602">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-603">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-603">1.7</span></span>|
|[<span data-ttu-id="00e05-604">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-605">ReadItem</span></span>|
|[<span data-ttu-id="00e05-606">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-607">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-608">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="00e05-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="00e05-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="00e05-610">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="00e05-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="00e05-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="00e05-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-613">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-613">Read mode</span></span>

<span data-ttu-id="00e05-614">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="00e05-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-615">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-615">Compose mode</span></span>

<span data-ttu-id="00e05-616">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="00e05-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="00e05-617">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="00e05-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-618">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-618">Type:</span></span>

*   <span data-ttu-id="00e05-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="00e05-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-620">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-620">Requirements</span></span>

|<span data-ttu-id="00e05-621">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-621">Requirement</span></span>|<span data-ttu-id="00e05-622">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-623">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-624">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-624">1.0</span></span>|
|[<span data-ttu-id="00e05-625">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-626">ReadItem</span></span>|
|[<span data-ttu-id="00e05-627">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-628">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-629">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-629">Example</span></span>

<span data-ttu-id="00e05-630">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="00e05-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="00e05-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00e05-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="00e05-632">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="00e05-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="00e05-633">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="00e05-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-634">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-634">Read mode</span></span>

<span data-ttu-id="00e05-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="00e05-637">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-637">Compose mode</span></span>

<span data-ttu-id="00e05-638">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="00e05-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00e05-639">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-639">Type:</span></span>

*   <span data-ttu-id="00e05-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00e05-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-641">Requirements</span></span>

|<span data-ttu-id="00e05-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-642">Requirement</span></span>|<span data-ttu-id="00e05-643">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-644">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-645">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-645">1.0</span></span>|
|[<span data-ttu-id="00e05-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-647">ReadItem</span></span>|
|[<span data-ttu-id="00e05-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-649">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="00e05-650">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="00e05-651">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="00e05-652">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00e05-653">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-653">Read mode</span></span>

<span data-ttu-id="00e05-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="00e05-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="00e05-656">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="00e05-656">Compose mode</span></span>

<span data-ttu-id="00e05-657">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-657">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="00e05-658">Tipo:</span><span class="sxs-lookup"><span data-stu-id="00e05-658">Type:</span></span>

*   <span data-ttu-id="00e05-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00e05-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-660">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-660">Requirements</span></span>

|<span data-ttu-id="00e05-661">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-661">Requirement</span></span>|<span data-ttu-id="00e05-662">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-663">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-664">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-664">1.0</span></span>|
|[<span data-ttu-id="00e05-665">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-666">ReadItem</span></span>|
|[<span data-ttu-id="00e05-667">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-668">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-669">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="00e05-670">Métodos</span><span class="sxs-lookup"><span data-stu-id="00e05-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="00e05-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00e05-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00e05-672">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="00e05-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="00e05-673">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="00e05-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="00e05-674">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="00e05-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-675">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-675">Parameters:</span></span>
|<span data-ttu-id="00e05-676">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-676">Name</span></span>|<span data-ttu-id="00e05-677">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-677">Type</span></span>|<span data-ttu-id="00e05-678">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-678">Attributes</span></span>|<span data-ttu-id="00e05-679">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="00e05-680">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-680">String</span></span>||<span data-ttu-id="00e05-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="00e05-683">String</span><span class="sxs-lookup"><span data-stu-id="00e05-683">String</span></span>||<span data-ttu-id="00e05-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="00e05-686">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-686">Object</span></span>|<span data-ttu-id="00e05-687">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-687">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-688">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-689">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-689">Object</span></span>|<span data-ttu-id="00e05-690">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-690">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-691">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="00e05-692">Booliano</span><span class="sxs-lookup"><span data-stu-id="00e05-692">Boolean</span></span>|<span data-ttu-id="00e05-693">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-693">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-694">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="00e05-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="00e05-695">function</span><span class="sxs-lookup"><span data-stu-id="00e05-695">function</span></span>|<span data-ttu-id="00e05-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-696">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-697">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00e05-698">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00e05-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00e05-699">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="00e05-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00e05-700">Erros</span><span class="sxs-lookup"><span data-stu-id="00e05-700">Errors</span></span>

|<span data-ttu-id="00e05-701">Código de erro</span><span class="sxs-lookup"><span data-stu-id="00e05-701">Error code</span></span>|<span data-ttu-id="00e05-702">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="00e05-703">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="00e05-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="00e05-704">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="00e05-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="00e05-705">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="00e05-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-706">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-706">Requirements</span></span>

|<span data-ttu-id="00e05-707">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-707">Requirement</span></span>|<span data-ttu-id="00e05-708">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-709">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-710">1.1</span><span class="sxs-lookup"><span data-stu-id="00e05-710">1.1</span></span>|
|[<span data-ttu-id="00e05-711">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-713">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-714">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00e05-715">Exemplos</span><span class="sxs-lookup"><span data-stu-id="00e05-715">Examples</span></span>

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

<span data-ttu-id="00e05-716">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="00e05-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00e05-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="00e05-718">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="00e05-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="00e05-719">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="00e05-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-720">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-720">Parameters:</span></span>

| <span data-ttu-id="00e05-721">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-721">Name</span></span> | <span data-ttu-id="00e05-722">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-722">Type</span></span> | <span data-ttu-id="00e05-723">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-723">Attributes</span></span> | <span data-ttu-id="00e05-724">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="00e05-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="00e05-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="00e05-726">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="00e05-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="00e05-727">Função</span><span class="sxs-lookup"><span data-stu-id="00e05-727">Function</span></span> || <span data-ttu-id="00e05-p136">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="00e05-731">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-731">Object</span></span> | <span data-ttu-id="00e05-732">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-732">&lt;optional&gt;</span></span> | <span data-ttu-id="00e05-733">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="00e05-734">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-734">Object</span></span> | <span data-ttu-id="00e05-735">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-735">&lt;optional&gt;</span></span> | <span data-ttu-id="00e05-736">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="00e05-737">function</span><span class="sxs-lookup"><span data-stu-id="00e05-737">function</span></span>| <span data-ttu-id="00e05-738">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-738">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-739">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-740">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-740">Requirements</span></span>

|<span data-ttu-id="00e05-741">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-741">Requirement</span></span>| <span data-ttu-id="00e05-742">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-743">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00e05-744">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-744">1.7</span></span> |
|[<span data-ttu-id="00e05-745">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00e05-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-746">ReadItem</span></span> |
|[<span data-ttu-id="00e05-747">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00e05-748">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="00e05-749">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="00e05-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00e05-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00e05-751">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="00e05-p137">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="00e05-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="00e05-755">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="00e05-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="00e05-756">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="00e05-756">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-757">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-757">Parameters:</span></span>

|<span data-ttu-id="00e05-758">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-758">Name</span></span>|<span data-ttu-id="00e05-759">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-759">Type</span></span>|<span data-ttu-id="00e05-760">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-760">Attributes</span></span>|<span data-ttu-id="00e05-761">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="00e05-762">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-762">String</span></span>||<span data-ttu-id="00e05-p138">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="00e05-765">String</span><span class="sxs-lookup"><span data-stu-id="00e05-765">String</span></span>||<span data-ttu-id="00e05-p139">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="00e05-768">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-768">Object</span></span>|<span data-ttu-id="00e05-769">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-769">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-770">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-771">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-771">Object</span></span>|<span data-ttu-id="00e05-772">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-772">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-773">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="00e05-774">function</span><span class="sxs-lookup"><span data-stu-id="00e05-774">function</span></span>|<span data-ttu-id="00e05-775">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-775">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-776">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00e05-777">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00e05-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00e05-778">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="00e05-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00e05-779">Erros</span><span class="sxs-lookup"><span data-stu-id="00e05-779">Errors</span></span>

|<span data-ttu-id="00e05-780">Código de erro</span><span class="sxs-lookup"><span data-stu-id="00e05-780">Error code</span></span>|<span data-ttu-id="00e05-781">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="00e05-782">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="00e05-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-783">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-783">Requirements</span></span>

|<span data-ttu-id="00e05-784">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-784">Requirement</span></span>|<span data-ttu-id="00e05-785">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-786">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-787">1.1</span><span class="sxs-lookup"><span data-stu-id="00e05-787">1.1</span></span>|
|[<span data-ttu-id="00e05-788">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-790">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-791">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-792">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-792">Example</span></span>

<span data-ttu-id="00e05-793">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="00e05-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="00e05-794">close()</span><span class="sxs-lookup"><span data-stu-id="00e05-794">close()</span></span>

<span data-ttu-id="00e05-795">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="00e05-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="00e05-p140">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="00e05-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-798">No Outlook na web, se o item é um compromisso e ele tenha sido salva anteriormente usando `saveAsync`, o usuário é solicitado a salvar, descartar ou Cancelar, mesmo se nenhuma alteração ocorreram desde que o item foi último salvo.</span><span class="sxs-lookup"><span data-stu-id="00e05-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="00e05-799">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="00e05-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-800">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-800">Requirements</span></span>

|<span data-ttu-id="00e05-801">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-801">Requirement</span></span>|<span data-ttu-id="00e05-802">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-803">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-804">1.3</span><span class="sxs-lookup"><span data-stu-id="00e05-804">1.3</span></span>|
|[<span data-ttu-id="00e05-805">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-806">Restrito</span><span class="sxs-lookup"><span data-stu-id="00e05-806">Restricted</span></span>|
|[<span data-ttu-id="00e05-807">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-808">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="00e05-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="00e05-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="00e05-810">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="00e05-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-811">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-811">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-812">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="00e05-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00e05-813">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="00e05-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="00e05-p141">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="00e05-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-817">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-817">Parameters:</span></span>

|<span data-ttu-id="00e05-818">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-818">Name</span></span>|<span data-ttu-id="00e05-819">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-819">Type</span></span>|<span data-ttu-id="00e05-820">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-820">Attributes</span></span>|<span data-ttu-id="00e05-821">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="00e05-822">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00e05-822">String &#124; Object</span></span>||<span data-ttu-id="00e05-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="00e05-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00e05-825">**OU**</span><span class="sxs-lookup"><span data-stu-id="00e05-825">**OR**</span></span><br/><span data-ttu-id="00e05-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="00e05-828">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-828">String</span></span>|<span data-ttu-id="00e05-829">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-829">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="00e05-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="00e05-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="00e05-833">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-833">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-834">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="00e05-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="00e05-835">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-835">String</span></span>||<span data-ttu-id="00e05-p145">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="00e05-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="00e05-838">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-838">String</span></span>||<span data-ttu-id="00e05-839">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="00e05-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="00e05-840">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-840">String</span></span>||<span data-ttu-id="00e05-p146">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="00e05-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="00e05-843">Boolean</span><span class="sxs-lookup"><span data-stu-id="00e05-843">Boolean</span></span>||<span data-ttu-id="00e05-p147">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="00e05-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="00e05-846">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-846">String</span></span>||<span data-ttu-id="00e05-p148">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="00e05-850">function</span><span class="sxs-lookup"><span data-stu-id="00e05-850">function</span></span>|<span data-ttu-id="00e05-851">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-851">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-852">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-853">Requirements</span></span>

|<span data-ttu-id="00e05-854">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-854">Requirement</span></span>|<span data-ttu-id="00e05-855">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-856">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-857">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-857">1.0</span></span>|
|[<span data-ttu-id="00e05-858">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-859">ReadItem</span></span>|
|[<span data-ttu-id="00e05-860">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-861">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00e05-862">Exemplos</span><span class="sxs-lookup"><span data-stu-id="00e05-862">Examples</span></span>

<span data-ttu-id="00e05-863">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="00e05-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="00e05-864">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="00e05-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="00e05-865">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="00e05-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00e05-866">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="00e05-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00e05-867">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="00e05-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00e05-868">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="00e05-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="00e05-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="00e05-870">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="00e05-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-871">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-871">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-872">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="00e05-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00e05-873">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="00e05-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="00e05-p149">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="00e05-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-877">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-877">Parameters:</span></span>

|<span data-ttu-id="00e05-878">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-878">Name</span></span>|<span data-ttu-id="00e05-879">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-879">Type</span></span>|<span data-ttu-id="00e05-880">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-880">Attributes</span></span>|<span data-ttu-id="00e05-881">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="00e05-882">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00e05-882">String &#124; Object</span></span>||<span data-ttu-id="00e05-p150">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="00e05-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00e05-885">**OU**</span><span class="sxs-lookup"><span data-stu-id="00e05-885">**OR**</span></span><br/><span data-ttu-id="00e05-p151">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="00e05-888">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-888">String</span></span>|<span data-ttu-id="00e05-889">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-889">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="00e05-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="00e05-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="00e05-893">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-893">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-894">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="00e05-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="00e05-895">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-895">String</span></span>||<span data-ttu-id="00e05-p153">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="00e05-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="00e05-898">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-898">String</span></span>||<span data-ttu-id="00e05-899">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="00e05-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="00e05-900">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-900">String</span></span>||<span data-ttu-id="00e05-p154">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="00e05-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="00e05-903">Boolean</span><span class="sxs-lookup"><span data-stu-id="00e05-903">Boolean</span></span>||<span data-ttu-id="00e05-p155">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="00e05-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="00e05-906">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-906">String</span></span>||<span data-ttu-id="00e05-p156">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="00e05-910">function</span><span class="sxs-lookup"><span data-stu-id="00e05-910">function</span></span>|<span data-ttu-id="00e05-911">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-911">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-912">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-913">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-913">Requirements</span></span>

|<span data-ttu-id="00e05-914">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-914">Requirement</span></span>|<span data-ttu-id="00e05-915">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-916">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-917">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-917">1.0</span></span>|
|[<span data-ttu-id="00e05-918">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-919">ReadItem</span></span>|
|[<span data-ttu-id="00e05-920">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-921">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00e05-922">Exemplos</span><span class="sxs-lookup"><span data-stu-id="00e05-922">Examples</span></span>

<span data-ttu-id="00e05-923">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="00e05-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="00e05-924">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="00e05-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="00e05-925">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="00e05-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00e05-926">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="00e05-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00e05-927">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="00e05-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00e05-928">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="00e05-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="00e05-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="00e05-930">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="00e05-930">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-931">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-931">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-932">Requirements</span></span>

|<span data-ttu-id="00e05-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-933">Requirement</span></span>|<span data-ttu-id="00e05-934">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-935">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-936">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-936">1.0</span></span>|
|[<span data-ttu-id="00e05-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-938">ReadItem</span></span>|
|[<span data-ttu-id="00e05-939">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-940">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-941">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-941">Returns:</span></span>

<span data-ttu-id="00e05-942">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="00e05-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="00e05-943">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-943">Example</span></span>

<span data-ttu-id="00e05-944">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-944">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="00e05-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00e05-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00e05-946">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="00e05-946">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-947">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-947">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-948">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-948">Parameters:</span></span>

|<span data-ttu-id="00e05-949">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-949">Name</span></span>|<span data-ttu-id="00e05-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-950">Type</span></span>|<span data-ttu-id="00e05-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="00e05-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="00e05-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="00e05-953">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="00e05-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-954">Requirements</span></span>

|<span data-ttu-id="00e05-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-955">Requirement</span></span>|<span data-ttu-id="00e05-956">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-957">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-958">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-958">1.0</span></span>|
|[<span data-ttu-id="00e05-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-960">Restrito</span><span class="sxs-lookup"><span data-stu-id="00e05-960">Restricted</span></span>|
|[<span data-ttu-id="00e05-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-962">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-963">Returns:</span></span>

<span data-ttu-id="00e05-964">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="00e05-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="00e05-965">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="00e05-965">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="00e05-966">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="00e05-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="00e05-967">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="00e05-968">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="00e05-968">Value of `entityType`</span></span>|<span data-ttu-id="00e05-969">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="00e05-969">Type of objects in returned array</span></span>|<span data-ttu-id="00e05-970">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="00e05-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="00e05-971">String</span><span class="sxs-lookup"><span data-stu-id="00e05-971">String</span></span>|<span data-ttu-id="00e05-972">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="00e05-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="00e05-973">Contato</span><span class="sxs-lookup"><span data-stu-id="00e05-973">Contact</span></span>|<span data-ttu-id="00e05-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00e05-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="00e05-975">String</span><span class="sxs-lookup"><span data-stu-id="00e05-975">String</span></span>|<span data-ttu-id="00e05-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00e05-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="00e05-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="00e05-977">MeetingSuggestion</span></span>|<span data-ttu-id="00e05-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00e05-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="00e05-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="00e05-979">PhoneNumber</span></span>|<span data-ttu-id="00e05-980">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="00e05-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="00e05-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="00e05-981">TaskSuggestion</span></span>|<span data-ttu-id="00e05-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00e05-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="00e05-983">String</span><span class="sxs-lookup"><span data-stu-id="00e05-983">String</span></span>|<span data-ttu-id="00e05-984">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="00e05-984">**Restricted**</span></span>|

<span data-ttu-id="00e05-985">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00e05-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="00e05-986">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-986">Example</span></span>

<span data-ttu-id="00e05-987">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="00e05-987">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="00e05-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00e05-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00e05-989">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="00e05-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-990">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-990">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-991">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="00e05-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-992">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-992">Parameters:</span></span>

|<span data-ttu-id="00e05-993">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-993">Name</span></span>|<span data-ttu-id="00e05-994">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-994">Type</span></span>|<span data-ttu-id="00e05-995">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="00e05-996">String</span><span class="sxs-lookup"><span data-stu-id="00e05-996">String</span></span>|<span data-ttu-id="00e05-997">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="00e05-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-998">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-998">Requirements</span></span>

|<span data-ttu-id="00e05-999">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-999">Requirement</span></span>|<span data-ttu-id="00e05-1000">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1001">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-1002">1.0</span></span>|
|[<span data-ttu-id="00e05-1003">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1004">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1005">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1006">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1007">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1007">Returns:</span></span>

<span data-ttu-id="00e05-p158">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="00e05-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="00e05-1010">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00e05-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="00e05-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="00e05-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="00e05-1012">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="00e05-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1013">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-p159">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="00e05-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="00e05-1017">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="00e05-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="00e05-1018">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="00e05-p160">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="00e05-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-1022">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1022">Requirements</span></span>

|<span data-ttu-id="00e05-1023">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1023">Requirement</span></span>|<span data-ttu-id="00e05-1024">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1025">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-1026">1.0</span></span>|
|[<span data-ttu-id="00e05-1027">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1028">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1029">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1030">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1031">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1031">Returns:</span></span>

<span data-ttu-id="00e05-p161">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="00e05-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="00e05-1034">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="00e05-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00e05-1035">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00e05-1036">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1036">Example</span></span>

<span data-ttu-id="00e05-1037">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="00e05-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="00e05-1038">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="00e05-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="00e05-1039">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="00e05-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1040">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-1041">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="00e05-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="00e05-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="00e05-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1044">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1044">Parameters:</span></span>

|<span data-ttu-id="00e05-1045">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1045">Name</span></span>|<span data-ttu-id="00e05-1046">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1046">Type</span></span>|<span data-ttu-id="00e05-1047">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="00e05-1048">String</span><span class="sxs-lookup"><span data-stu-id="00e05-1048">String</span></span>|<span data-ttu-id="00e05-1049">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="00e05-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1050">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1050">Requirements</span></span>

|<span data-ttu-id="00e05-1051">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1051">Requirement</span></span>|<span data-ttu-id="00e05-1052">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1053">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-1054">1.0</span></span>|
|[<span data-ttu-id="00e05-1055">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1056">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1057">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1058">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1059">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1059">Returns:</span></span>

<span data-ttu-id="00e05-1060">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="00e05-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="00e05-1061">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="00e05-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00e05-1062">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="00e05-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00e05-1063">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="00e05-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="00e05-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="00e05-1065">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="00e05-p163">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1068">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1068">Parameters:</span></span>

|<span data-ttu-id="00e05-1069">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1069">Name</span></span>|<span data-ttu-id="00e05-1070">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1070">Type</span></span>|<span data-ttu-id="00e05-1071">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1071">Attributes</span></span>|<span data-ttu-id="00e05-1072">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="00e05-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00e05-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="00e05-p164">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="00e05-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="00e05-1077">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1077">Object</span></span>|<span data-ttu-id="00e05-1078">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1079">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-1080">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1080">Object</span></span>|<span data-ttu-id="00e05-1081">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1082">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="00e05-1083">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1083">function</span></span>||<span data-ttu-id="00e05-1084">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00e05-1085">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="00e05-1086">Para acessar a propriedade source que a seleção é proveniente, chamar `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1086">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1087">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1087">Requirements</span></span>

|<span data-ttu-id="00e05-1088">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1088">Requirement</span></span>|<span data-ttu-id="00e05-1089">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1090">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="00e05-1091">1.2</span></span>|
|[<span data-ttu-id="00e05-1092">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-1094">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1095">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1096">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1096">Returns:</span></span>

<span data-ttu-id="00e05-1097">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="00e05-1098">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="00e05-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00e05-1099">String</span><span class="sxs-lookup"><span data-stu-id="00e05-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00e05-1100">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1100">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="00e05-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="00e05-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="00e05-p166">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="00e05-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1104">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-1104">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-1105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1105">Requirements</span></span>

|<span data-ttu-id="00e05-1106">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1106">Requirement</span></span>|<span data-ttu-id="00e05-1107">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="00e05-1109">1.6</span></span>|
|[<span data-ttu-id="00e05-1110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1111">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1113">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1114">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1114">Returns:</span></span>

<span data-ttu-id="00e05-1115">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="00e05-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="00e05-1116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1116">Example</span></span>

<span data-ttu-id="00e05-1117">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="00e05-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="00e05-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="00e05-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="00e05-p167">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="00e05-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1121">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="00e05-1121">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00e05-p168">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="00e05-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="00e05-1125">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="00e05-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="00e05-1126">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="00e05-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="00e05-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00e05-1130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1130">Requirements</span></span>

|<span data-ttu-id="00e05-1131">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1131">Requirement</span></span>|<span data-ttu-id="00e05-1132">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="00e05-1134">1.6</span></span>|
|[<span data-ttu-id="00e05-1135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1136">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1138">Leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00e05-1139">Retorna:</span><span class="sxs-lookup"><span data-stu-id="00e05-1139">Returns:</span></span>

<span data-ttu-id="00e05-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="00e05-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="00e05-1142">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1142">Example</span></span>

<span data-ttu-id="00e05-1143">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="00e05-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="00e05-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="00e05-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="00e05-1145">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="00e05-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="00e05-p171">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="00e05-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1149">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1149">Parameters:</span></span>

|<span data-ttu-id="00e05-1150">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1150">Name</span></span>|<span data-ttu-id="00e05-1151">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1151">Type</span></span>|<span data-ttu-id="00e05-1152">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1152">Attributes</span></span>|<span data-ttu-id="00e05-1153">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="00e05-1154">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1154">function</span></span>||<span data-ttu-id="00e05-1155">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00e05-1156">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00e05-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="00e05-1157">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="00e05-1157">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="00e05-1158">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1158">Object</span></span>|<span data-ttu-id="00e05-1159">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1160">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1160">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="00e05-1161">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1162">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1162">Requirements</span></span>

|<span data-ttu-id="00e05-1163">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1163">Requirement</span></span>|<span data-ttu-id="00e05-1164">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1165">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="00e05-1166">1.0</span></span>|
|[<span data-ttu-id="00e05-1167">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1168">ReadItem</span></span>|
|[<span data-ttu-id="00e05-1169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1170">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-1171">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1171">Example</span></span>

<span data-ttu-id="00e05-p174">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="00e05-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="00e05-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00e05-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="00e05-1176">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="00e05-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="00e05-p175">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="00e05-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1181">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1181">Parameters:</span></span>

|<span data-ttu-id="00e05-1182">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1182">Name</span></span>|<span data-ttu-id="00e05-1183">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1183">Type</span></span>|<span data-ttu-id="00e05-1184">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1184">Attributes</span></span>|<span data-ttu-id="00e05-1185">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="00e05-1186">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="00e05-1186">String</span></span>||<span data-ttu-id="00e05-p176">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="00e05-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="00e05-1189">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1189">Object</span></span>|<span data-ttu-id="00e05-1190">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1191">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-1192">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1192">Object</span></span>|<span data-ttu-id="00e05-1193">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1194">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="00e05-1195">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1195">function</span></span>|<span data-ttu-id="00e05-1196">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1197">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00e05-1198">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="00e05-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00e05-1199">Erros</span><span class="sxs-lookup"><span data-stu-id="00e05-1199">Errors</span></span>

|<span data-ttu-id="00e05-1200">Código de erro</span><span class="sxs-lookup"><span data-stu-id="00e05-1200">Error code</span></span>|<span data-ttu-id="00e05-1201">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="00e05-1202">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="00e05-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1203">Requirements</span></span>

|<span data-ttu-id="00e05-1204">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1204">Requirement</span></span>|<span data-ttu-id="00e05-1205">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="00e05-1207">1.1</span></span>|
|[<span data-ttu-id="00e05-1208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-1210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1211">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-1212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1212">Example</span></span>

<span data-ttu-id="00e05-1213">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="00e05-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="00e05-1214">removeHandlerAsync (eventType, manipulador, [Opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="00e05-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="00e05-1215">Remove um manipulador de eventos para um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="00e05-1215">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="00e05-1216">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="00e05-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1217">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1217">Parameters:</span></span>

| <span data-ttu-id="00e05-1218">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1218">Name</span></span> | <span data-ttu-id="00e05-1219">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1219">Type</span></span> | <span data-ttu-id="00e05-1220">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1220">Attributes</span></span> | <span data-ttu-id="00e05-1221">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="00e05-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="00e05-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="00e05-1223">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="00e05-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="00e05-1224">Função</span><span class="sxs-lookup"><span data-stu-id="00e05-1224">Function</span></span> || <span data-ttu-id="00e05-p177">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="00e05-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1228">Object</span></span> | <span data-ttu-id="00e05-1229">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="00e05-1230">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="00e05-1231">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1231">Object</span></span> | <span data-ttu-id="00e05-1232">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="00e05-1233">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="00e05-1234">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1234">function</span></span>| <span data-ttu-id="00e05-1235">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1236">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1237">Requirements</span></span>

|<span data-ttu-id="00e05-1238">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1238">Requirement</span></span>| <span data-ttu-id="00e05-1239">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00e05-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="00e05-1241">1.7</span></span> |
|[<span data-ttu-id="00e05-1242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00e05-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1243">ReadItem</span></span> |
|[<span data-ttu-id="00e05-1244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00e05-1245">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="00e05-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="00e05-1246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="00e05-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="00e05-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="00e05-1248">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="00e05-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="00e05-p178">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="00e05-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1252">Se seu suplemento chama `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, pode levar algum tempo antes do item realmente é sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="00e05-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="00e05-1253">Até que o item é sincronizado, usando o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="00e05-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="00e05-p180">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="00e05-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="00e05-1257">Os seguintes clientes tenham comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="00e05-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="00e05-1258">Mac Outlook não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="00e05-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="00e05-1259">Chamar `saveAsync` em uma reunião no Outlook Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="00e05-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="00e05-1260">Outlook na web sempre envia um convite ou atualizar quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="00e05-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1261">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1261">Parameters:</span></span>

|<span data-ttu-id="00e05-1262">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1262">Name</span></span>|<span data-ttu-id="00e05-1263">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1263">Type</span></span>|<span data-ttu-id="00e05-1264">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1264">Attributes</span></span>|<span data-ttu-id="00e05-1265">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="00e05-1266">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1266">Object</span></span>|<span data-ttu-id="00e05-1267">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1268">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-1269">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1269">Object</span></span>|<span data-ttu-id="00e05-1270">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1271">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="00e05-1272">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1272">function</span></span>||<span data-ttu-id="00e05-1273">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00e05-1274">Em caso de sucesso, o identificador do item é fornecido no `asyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="00e05-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1275">Requirements</span></span>

|<span data-ttu-id="00e05-1276">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1276">Requirement</span></span>|<span data-ttu-id="00e05-1277">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="00e05-1279">1.3</span></span>|
|[<span data-ttu-id="00e05-1280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-1282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1283">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00e05-1284">Exemplos</span><span class="sxs-lookup"><span data-stu-id="00e05-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="00e05-p182">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="00e05-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="00e05-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="00e05-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="00e05-1288">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00e05-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="00e05-p183">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="00e05-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00e05-1292">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="00e05-1292">Parameters:</span></span>

|<span data-ttu-id="00e05-1293">Nome</span><span class="sxs-lookup"><span data-stu-id="00e05-1293">Name</span></span>|<span data-ttu-id="00e05-1294">Tipo</span><span class="sxs-lookup"><span data-stu-id="00e05-1294">Type</span></span>|<span data-ttu-id="00e05-1295">Atributos</span><span class="sxs-lookup"><span data-stu-id="00e05-1295">Attributes</span></span>|<span data-ttu-id="00e05-1296">Descrição</span><span class="sxs-lookup"><span data-stu-id="00e05-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="00e05-1297">String</span><span class="sxs-lookup"><span data-stu-id="00e05-1297">String</span></span>||<span data-ttu-id="00e05-p184">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="00e05-1301">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1301">Object</span></span>|<span data-ttu-id="00e05-1302">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1303">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="00e05-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="00e05-1304">Objeto</span><span class="sxs-lookup"><span data-stu-id="00e05-1304">Object</span></span>|<span data-ttu-id="00e05-1305">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-1306">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00e05-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="00e05-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00e05-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="00e05-1308">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="00e05-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="00e05-p185">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="00e05-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="00e05-p186">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="00e05-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="00e05-1313">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="00e05-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="00e05-1314">function</span><span class="sxs-lookup"><span data-stu-id="00e05-1314">function</span></span>||<span data-ttu-id="00e05-1315">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00e05-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00e05-1316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="00e05-1316">Requirements</span></span>

|<span data-ttu-id="00e05-1317">Requisito</span><span class="sxs-lookup"><span data-stu-id="00e05-1317">Requirement</span></span>|<span data-ttu-id="00e05-1318">Valor</span><span class="sxs-lookup"><span data-stu-id="00e05-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="00e05-1319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="00e05-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="00e05-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="00e05-1320">1.2</span></span>|
|[<span data-ttu-id="00e05-1321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="00e05-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="00e05-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00e05-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="00e05-1323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="00e05-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="00e05-1324">Composição</span><span class="sxs-lookup"><span data-stu-id="00e05-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00e05-1325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="00e05-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```