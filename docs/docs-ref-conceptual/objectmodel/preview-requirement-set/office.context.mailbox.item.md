
# <a name="item"></a><span data-ttu-id="8a6c0-101">item</span><span class="sxs-lookup"><span data-stu-id="8a6c0-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8a6c0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8a6c0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8a6c0-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-105">Requirements</span></span>

|<span data-ttu-id="8a6c0-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-106">Requirement</span></span>|<span data-ttu-id="8a6c0-107">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-109">1.0</span></span>|
|[<span data-ttu-id="8a6c0-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-111">Restricted</span></span>|
|[<span data-ttu-id="8a6c0-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8a6c0-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-114">Members and methods</span></span>

| <span data-ttu-id="8a6c0-115">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-115">Member</span></span> | <span data-ttu-id="8a6c0-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8a6c0-117">attachments</span><span class="sxs-lookup"><span data-stu-id="8a6c0-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8a6c0-118">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-118">Member</span></span> |
| [<span data-ttu-id="8a6c0-119">bcc</span><span class="sxs-lookup"><span data-stu-id="8a6c0-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8a6c0-120">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-120">Member</span></span> |
| [<span data-ttu-id="8a6c0-121">body</span><span class="sxs-lookup"><span data-stu-id="8a6c0-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="8a6c0-122">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-122">Member</span></span> |
| [<span data-ttu-id="8a6c0-123">cc</span><span class="sxs-lookup"><span data-stu-id="8a6c0-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8a6c0-124">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-124">Member</span></span> |
| [<span data-ttu-id="8a6c0-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="8a6c0-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8a6c0-126">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-126">Member</span></span> |
| [<span data-ttu-id="8a6c0-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8a6c0-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8a6c0-128">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-128">Member</span></span> |
| [<span data-ttu-id="8a6c0-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8a6c0-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8a6c0-130">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-130">Member</span></span> |
| [<span data-ttu-id="8a6c0-131">end</span><span class="sxs-lookup"><span data-stu-id="8a6c0-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8a6c0-132">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-132">Member</span></span> |
| [<span data-ttu-id="8a6c0-133">from</span><span class="sxs-lookup"><span data-stu-id="8a6c0-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="8a6c0-134">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-134">Member</span></span> |
| [<span data-ttu-id="8a6c0-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8a6c0-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8a6c0-136">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-136">Member</span></span> |
| [<span data-ttu-id="8a6c0-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="8a6c0-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8a6c0-138">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-138">Member</span></span> |
| [<span data-ttu-id="8a6c0-139">itemId</span><span class="sxs-lookup"><span data-stu-id="8a6c0-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8a6c0-140">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-140">Member</span></span> |
| [<span data-ttu-id="8a6c0-141">itemType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="8a6c0-142">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-142">Member</span></span> |
| [<span data-ttu-id="8a6c0-143">location</span><span class="sxs-lookup"><span data-stu-id="8a6c0-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="8a6c0-144">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-144">Member</span></span> |
| [<span data-ttu-id="8a6c0-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8a6c0-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8a6c0-146">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-146">Member</span></span> |
| [<span data-ttu-id="8a6c0-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8a6c0-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="8a6c0-148">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-148">Member</span></span> |
| [<span data-ttu-id="8a6c0-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8a6c0-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8a6c0-150">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-150">Member</span></span> |
| [<span data-ttu-id="8a6c0-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8a6c0-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="8a6c0-152">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-152">Member</span></span> |
| [<span data-ttu-id="8a6c0-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="8a6c0-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="8a6c0-154">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-154">Member</span></span> |
| [<span data-ttu-id="8a6c0-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8a6c0-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8a6c0-156">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-156">Member</span></span> |
| [<span data-ttu-id="8a6c0-157">sender</span><span class="sxs-lookup"><span data-stu-id="8a6c0-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="8a6c0-158">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-158">Member</span></span> |
| [<span data-ttu-id="8a6c0-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="8a6c0-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="8a6c0-160">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-160">Member</span></span> |
| [<span data-ttu-id="8a6c0-161">start</span><span class="sxs-lookup"><span data-stu-id="8a6c0-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8a6c0-162">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-162">Member</span></span> |
| [<span data-ttu-id="8a6c0-163">subject</span><span class="sxs-lookup"><span data-stu-id="8a6c0-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="8a6c0-164">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-164">Member</span></span> |
| [<span data-ttu-id="8a6c0-165">to</span><span class="sxs-lookup"><span data-stu-id="8a6c0-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8a6c0-166">Membro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-166">Member</span></span> |
| [<span data-ttu-id="8a6c0-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8a6c0-168">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-168">Method</span></span> |
| [<span data-ttu-id="8a6c0-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8a6c0-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="8a6c0-170">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-170">Method</span></span> |
| [<span data-ttu-id="8a6c0-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8a6c0-172">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-172">Method</span></span> |
| [<span data-ttu-id="8a6c0-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8a6c0-174">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-174">Method</span></span> |
| [<span data-ttu-id="8a6c0-175">close</span><span class="sxs-lookup"><span data-stu-id="8a6c0-175">close</span></span>](#close) | <span data-ttu-id="8a6c0-176">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-176">Method</span></span> |
| [<span data-ttu-id="8a6c0-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8a6c0-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8a6c0-178">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-178">Method</span></span> |
| [<span data-ttu-id="8a6c0-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8a6c0-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8a6c0-180">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-180">Method</span></span> |
| [<span data-ttu-id="8a6c0-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="8a6c0-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8a6c0-182">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-182">Method</span></span> |
| [<span data-ttu-id="8a6c0-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8a6c0-184">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-184">Method</span></span> |
| [<span data-ttu-id="8a6c0-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8a6c0-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8a6c0-186">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-186">Method</span></span> |
| [<span data-ttu-id="8a6c0-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="8a6c0-188">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-188">Method</span></span> |
| [<span data-ttu-id="8a6c0-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8a6c0-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8a6c0-190">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-190">Method</span></span> |
| [<span data-ttu-id="8a6c0-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8a6c0-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8a6c0-192">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-192">Method</span></span> |
| [<span data-ttu-id="8a6c0-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8a6c0-194">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-194">Method</span></span> |
| [<span data-ttu-id="8a6c0-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8a6c0-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8a6c0-196">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-196">Method</span></span> |
| [<span data-ttu-id="8a6c0-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8a6c0-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8a6c0-198">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-198">Method</span></span> |
| [<span data-ttu-id="8a6c0-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="8a6c0-200">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-200">Method</span></span> |
| [<span data-ttu-id="8a6c0-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8a6c0-202">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-202">Method</span></span> |
| [<span data-ttu-id="8a6c0-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8a6c0-204">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-204">Method</span></span> |
| [<span data-ttu-id="8a6c0-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8a6c0-206">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-206">Method</span></span> |
| [<span data-ttu-id="8a6c0-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8a6c0-208">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-208">Method</span></span> |
| [<span data-ttu-id="8a6c0-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8a6c0-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8a6c0-210">Método</span><span class="sxs-lookup"><span data-stu-id="8a6c0-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8a6c0-211">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-211">Example</span></span>

<span data-ttu-id="8a6c0-212">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8a6c0-213">Membros</span><span class="sxs-lookup"><span data-stu-id="8a6c0-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8a6c0-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8a6c0-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8a6c0-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-217">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8a6c0-218">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-219">Type:</span></span>

*   <span data-ttu-id="8a6c0-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8a6c0-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-221">Requirements</span></span>

|<span data-ttu-id="8a6c0-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-222">Requirement</span></span>|<span data-ttu-id="8a6c0-223">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-225">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-225">1.0</span></span>|
|[<span data-ttu-id="8a6c0-226">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-227">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-228">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-229">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-230">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-230">Example</span></span>

<span data-ttu-id="8a6c0-231">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8a6c0-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8a6c0-233">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8a6c0-234">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-235">Type:</span></span>

*   [<span data-ttu-id="8a6c0-236">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8a6c0-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8a6c0-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-237">Requirements</span></span>

|<span data-ttu-id="8a6c0-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-238">Requirement</span></span>|<span data-ttu-id="8a6c0-239">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-241">1.1</span><span class="sxs-lookup"><span data-stu-id="8a6c0-241">1.1</span></span>|
|[<span data-ttu-id="8a6c0-242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-243">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-245">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="8a6c0-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="8a6c0-248">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-249">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-249">Type:</span></span>

*   [<span data-ttu-id="8a6c0-250">Corpo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="8a6c0-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-251">Requirements</span></span>

|<span data-ttu-id="8a6c0-252">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-252">Requirement</span></span>|<span data-ttu-id="8a6c0-253">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-254">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-255">1.1</span><span class="sxs-lookup"><span data-stu-id="8a6c0-255">1.1</span></span>|
|[<span data-ttu-id="8a6c0-256">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-257">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-258">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-259">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8a6c0-260">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8a6c0-261">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8a6c0-262">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-263">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-263">Read mode</span></span>

<span data-ttu-id="8a6c0-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-266">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-266">Compose mode</span></span>

<span data-ttu-id="8a6c0-267">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-268">Type:</span></span>

*   <span data-ttu-id="8a6c0-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-270">Requirements</span></span>

|<span data-ttu-id="8a6c0-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-271">Requirement</span></span>|<span data-ttu-id="8a6c0-272">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-274">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-274">1.0</span></span>|
|[<span data-ttu-id="8a6c0-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-276">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-278">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8a6c0-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="8a6c0-281">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8a6c0-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8a6c0-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-286">Type:</span></span>

*   <span data-ttu-id="8a6c0-287">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-288">Requirements</span></span>

|<span data-ttu-id="8a6c0-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-289">Requirement</span></span>|<span data-ttu-id="8a6c0-290">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-292">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-292">1.0</span></span>|
|[<span data-ttu-id="8a6c0-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-294">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-296">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8a6c0-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8a6c0-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="8a6c0-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-300">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-300">Type:</span></span>

*   <span data-ttu-id="8a6c0-301">Data</span><span class="sxs-lookup"><span data-stu-id="8a6c0-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-302">Requirements</span></span>

|<span data-ttu-id="8a6c0-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-303">Requirement</span></span>|<span data-ttu-id="8a6c0-304">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-306">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-306">1.0</span></span>|
|[<span data-ttu-id="8a6c0-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-308">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-309">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-310">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-311">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8a6c0-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8a6c0-312">dateTimeModified :Date</span></span>

<span data-ttu-id="8a6c0-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-315">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-316">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-316">Type:</span></span>

*   <span data-ttu-id="8a6c0-317">Data</span><span class="sxs-lookup"><span data-stu-id="8a6c0-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-318">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-318">Requirements</span></span>

|<span data-ttu-id="8a6c0-319">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-319">Requirement</span></span>|<span data-ttu-id="8a6c0-320">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-321">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-322">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-322">1.0</span></span>|
|[<span data-ttu-id="8a6c0-323">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-324">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-325">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-326">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-327">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8a6c0-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8a6c0-329">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8a6c0-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-332">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-332">Read mode</span></span>

<span data-ttu-id="8a6c0-333">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-334">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-334">Compose mode</span></span>

<span data-ttu-id="8a6c0-335">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8a6c0-336">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-337">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-337">Type:</span></span>

*   <span data-ttu-id="8a6c0-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-339">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-339">Requirements</span></span>

|<span data-ttu-id="8a6c0-340">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-340">Requirement</span></span>|<span data-ttu-id="8a6c0-341">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-342">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-343">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-343">1.0</span></span>|
|[<span data-ttu-id="8a6c0-344">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-345">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-346">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-347">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-348">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-348">Example</span></span>

<span data-ttu-id="8a6c0-349">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="8a6c0-350">de:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="8a6c0-351">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="8a6c0-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-354">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-355">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-355">Read mode</span></span>

<span data-ttu-id="8a6c0-356">O `from` propriedade retornará uma `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-357">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-357">Compose mode</span></span>

<span data-ttu-id="8a6c0-358">O `from` propriedade retornará uma `From` objeto que fornece um método para obter o de valor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a6c0-359">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-359">Type:</span></span>

*   <span data-ttu-id="8a6c0-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-361">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-361">Requirements</span></span>

|<span data-ttu-id="8a6c0-362">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8a6c0-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-364">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-364">1.0</span></span>|<span data-ttu-id="8a6c0-365">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-365">1.7</span></span>|
|[<span data-ttu-id="8a6c0-366">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-367">ReadItem</span></span>|<span data-ttu-id="8a6c0-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-370">Read</span><span class="sxs-lookup"><span data-stu-id="8a6c0-370">Read</span></span>|<span data-ttu-id="8a6c0-371">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8a6c0-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-372">internetMessageId :String</span></span>

<span data-ttu-id="8a6c0-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-375">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-375">Type:</span></span>

*   <span data-ttu-id="8a6c0-376">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-377">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-377">Requirements</span></span>

|<span data-ttu-id="8a6c0-378">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-378">Requirement</span></span>|<span data-ttu-id="8a6c0-379">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-380">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-381">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-381">1.0</span></span>|
|[<span data-ttu-id="8a6c0-382">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-383">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-384">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-385">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-386">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8a6c0-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-387">itemClass :String</span></span>

<span data-ttu-id="8a6c0-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8a6c0-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="8a6c0-392">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-392">Type</span></span>|<span data-ttu-id="8a6c0-393">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-393">Description</span></span>|<span data-ttu-id="8a6c0-394">classe de item</span><span class="sxs-lookup"><span data-stu-id="8a6c0-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="8a6c0-395">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="8a6c0-395">Appointment items</span></span>|<span data-ttu-id="8a6c0-396">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="8a6c0-397">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-397">Message items</span></span>|<span data-ttu-id="8a6c0-398">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="8a6c0-399">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-400">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-400">Type:</span></span>

*   <span data-ttu-id="8a6c0-401">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-402">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-402">Requirements</span></span>

|<span data-ttu-id="8a6c0-403">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-403">Requirement</span></span>|<span data-ttu-id="8a6c0-404">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-405">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-406">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-406">1.0</span></span>|
|[<span data-ttu-id="8a6c0-407">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-408">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-409">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-410">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-411">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8a6c0-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-412">(nullable) itemId :String</span></span>

<span data-ttu-id="8a6c0-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-415">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8a6c0-416">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8a6c0-417">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8a6c0-418">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8a6c0-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-421">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-421">Type:</span></span>

*   <span data-ttu-id="8a6c0-422">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-423">Requirements</span></span>

|<span data-ttu-id="8a6c0-424">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-424">Requirement</span></span>|<span data-ttu-id="8a6c0-425">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-426">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-427">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-427">1.0</span></span>|
|[<span data-ttu-id="8a6c0-428">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-429">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-430">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-431">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-432">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-432">Example</span></span>

<span data-ttu-id="8a6c0-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="8a6c0-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8a6c0-436">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8a6c0-437">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-438">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-438">Type:</span></span>

*   [<span data-ttu-id="8a6c0-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8a6c0-440">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-440">Requirements</span></span>

|<span data-ttu-id="8a6c0-441">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-441">Requirement</span></span>|<span data-ttu-id="8a6c0-442">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-443">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-444">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-444">1.0</span></span>|
|[<span data-ttu-id="8a6c0-445">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-446">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-447">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-448">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-449">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="8a6c0-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="8a6c0-451">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-452">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-452">Read mode</span></span>

<span data-ttu-id="8a6c0-453">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-454">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-454">Compose mode</span></span>

<span data-ttu-id="8a6c0-455">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-456">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-456">Type:</span></span>

*   <span data-ttu-id="8a6c0-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-458">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-458">Requirements</span></span>

|<span data-ttu-id="8a6c0-459">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-459">Requirement</span></span>|<span data-ttu-id="8a6c0-460">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-461">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-462">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-462">1.0</span></span>|
|[<span data-ttu-id="8a6c0-463">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-464">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-465">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-466">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-467">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8a6c0-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-468">normalizedSubject :String</span></span>

<span data-ttu-id="8a6c0-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8a6c0-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-473">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-473">Type:</span></span>

*   <span data-ttu-id="8a6c0-474">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-475">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-475">Requirements</span></span>

|<span data-ttu-id="8a6c0-476">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-476">Requirement</span></span>|<span data-ttu-id="8a6c0-477">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-478">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-479">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-479">1.0</span></span>|
|[<span data-ttu-id="8a6c0-480">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-481">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-482">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-483">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-484">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="8a6c0-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="8a6c0-486">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-487">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-487">Type:</span></span>

*   [<span data-ttu-id="8a6c0-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8a6c0-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8a6c0-489">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-489">Requirements</span></span>

|<span data-ttu-id="8a6c0-490">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-490">Requirement</span></span>|<span data-ttu-id="8a6c0-491">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-492">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-493">1.3</span><span class="sxs-lookup"><span data-stu-id="8a6c0-493">1.3</span></span>|
|[<span data-ttu-id="8a6c0-494">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-495">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-496">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-497">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8a6c0-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8a6c0-499">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8a6c0-500">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-501">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-501">Read mode</span></span>

<span data-ttu-id="8a6c0-502">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-503">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-503">Compose mode</span></span>

<span data-ttu-id="8a6c0-504">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-505">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-505">Type:</span></span>

*   <span data-ttu-id="8a6c0-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-507">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-507">Requirements</span></span>

|<span data-ttu-id="8a6c0-508">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-508">Requirement</span></span>|<span data-ttu-id="8a6c0-509">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-510">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-511">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-511">1.0</span></span>|
|[<span data-ttu-id="8a6c0-512">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-513">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-514">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-515">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-516">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="8a6c0-517">Organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="8a6c0-518">Obtém o endereço de email do organizador da reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-519">Read mode</span></span>

<span data-ttu-id="8a6c0-520">O `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-521">Compose mode</span></span>

<span data-ttu-id="8a6c0-522">O `organizer` propriedade retorna um objeto do [Organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-523">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-523">Type:</span></span>

*   <span data-ttu-id="8a6c0-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-525">Requirements</span></span>

|<span data-ttu-id="8a6c0-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8a6c0-527">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-528">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-528">1.0</span></span>|<span data-ttu-id="8a6c0-529">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-529">1.7</span></span>|
|[<span data-ttu-id="8a6c0-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-531">ReadItem</span></span>|<span data-ttu-id="8a6c0-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-533">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-534">Read</span><span class="sxs-lookup"><span data-stu-id="8a6c0-534">Read</span></span>|<span data-ttu-id="8a6c0-535">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-536">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="8a6c0-537">Recorrência (anulável):[Recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="8a6c0-538">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="8a6c0-539">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="8a6c0-540">Leitura e redação modos para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="8a6c0-541">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="8a6c0-542">O `recurrence` propriedade retorna um objeto de [Recorrência](/javascript/api/outlook/office.recurrence) para solicitações de reuniões ou compromissos recorrentes, se um item é uma série ou uma instância de uma série.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="8a6c0-543">`null`é retornada para um único compromissos e solicitações de reunião de compromissos de única.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="8a6c0-544">`undefined`é retornado para mensagens que não fazem solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="8a6c0-545">Observação: Solicitações de reunião tem um `itemClass` valor IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="8a6c0-546">Observação: Se o objeto de recorrência for `null`, isto indica que o objeto é um compromisso único ou uma solicitação de reunião de um único compromisso e não fizer parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-547">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-547">Type:</span></span>

* [<span data-ttu-id="8a6c0-548">Recorrência</span><span class="sxs-lookup"><span data-stu-id="8a6c0-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="8a6c0-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-549">Requirement</span></span>|<span data-ttu-id="8a6c0-550">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-552">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-552">1.7</span></span>|
|[<span data-ttu-id="8a6c0-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-554">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-556">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8a6c0-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8a6c0-558">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8a6c0-559">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-560">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-560">Read mode</span></span>

<span data-ttu-id="8a6c0-561">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-562">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-562">Compose mode</span></span>

<span data-ttu-id="8a6c0-563">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-564">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-564">Type:</span></span>

*   <span data-ttu-id="8a6c0-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-566">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-566">Requirements</span></span>

|<span data-ttu-id="8a6c0-567">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-567">Requirement</span></span>|<span data-ttu-id="8a6c0-568">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-569">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-570">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-570">1.0</span></span>|
|[<span data-ttu-id="8a6c0-571">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-572">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-573">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-574">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-575">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="8a6c0-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="8a6c0-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8a6c0-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-581">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `sender` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-582">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-582">Type:</span></span>

*   [<span data-ttu-id="8a6c0-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8a6c0-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8a6c0-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-584">Requirements</span></span>

|<span data-ttu-id="8a6c0-585">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-585">Requirement</span></span>|<span data-ttu-id="8a6c0-586">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-588">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-588">1.0</span></span>|
|[<span data-ttu-id="8a6c0-589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-590">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-592">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="8a6c0-594">seriesId (anulável): cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="8a6c0-595">Obtém a identificação da série que uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="8a6c0-596">No OWA e no Outlook, o `seriesId` retornará a identificação de serviços Web do Exchange (EWS) do item pai (série) que este item pertence.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="8a6c0-597">No entanto, em iOS e Android, o `seriesId` retornará a identificação do restante do item pai.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-598">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8a6c0-599">O `seriesId` propriedade não é idêntica às IDs do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="8a6c0-600">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8a6c0-601">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="8a6c0-602">O `seriesId` propriedade retornará `null` para itens que não têm itens pai como único compromissos, itens de série, ou solicitações de reunião e retorna `undefined` para todos os itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-603">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-603">Type:</span></span>

* <span data-ttu-id="8a6c0-604">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-605">Requirements</span></span>

|<span data-ttu-id="8a6c0-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-606">Requirement</span></span>|<span data-ttu-id="8a6c0-607">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-609">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-609">1.7</span></span>|
|[<span data-ttu-id="8a6c0-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-611">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-613">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8a6c0-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8a6c0-616">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8a6c0-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-619">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-619">Read mode</span></span>

<span data-ttu-id="8a6c0-620">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-621">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-621">Compose mode</span></span>

<span data-ttu-id="8a6c0-622">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8a6c0-623">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-624">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-624">Type:</span></span>

*   <span data-ttu-id="8a6c0-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-626">Requirements</span></span>

|<span data-ttu-id="8a6c0-627">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-627">Requirement</span></span>|<span data-ttu-id="8a6c0-628">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-629">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-630">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-630">1.0</span></span>|
|[<span data-ttu-id="8a6c0-631">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-632">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-633">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-634">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-635">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-635">Example</span></span>

<span data-ttu-id="8a6c0-636">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="8a6c0-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="8a6c0-638">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8a6c0-639">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-640">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-640">Read mode</span></span>

<span data-ttu-id="8a6c0-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-643">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-643">Compose mode</span></span>

<span data-ttu-id="8a6c0-644">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a6c0-645">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-645">Type:</span></span>

*   <span data-ttu-id="8a6c0-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-647">Requirements</span></span>

|<span data-ttu-id="8a6c0-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-648">Requirement</span></span>|<span data-ttu-id="8a6c0-649">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-651">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-651">1.0</span></span>|
|[<span data-ttu-id="8a6c0-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-653">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-655">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8a6c0-656">para: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8a6c0-657">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8a6c0-658">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a6c0-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-659">Read mode</span></span>

<span data-ttu-id="8a6c0-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8a6c0-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8a6c0-662">Compose mode</span></span>

<span data-ttu-id="8a6c0-663">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8a6c0-664">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-664">Type:</span></span>

*   <span data-ttu-id="8a6c0-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-666">Requirements</span></span>

|<span data-ttu-id="8a6c0-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-667">Requirement</span></span>|<span data-ttu-id="8a6c0-668">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-670">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-670">1.0</span></span>|
|[<span data-ttu-id="8a6c0-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-672">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-674">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-675">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8a6c0-676">Métodos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8a6c0-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-678">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8a6c0-679">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8a6c0-680">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-681">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-681">Parameters:</span></span>
|<span data-ttu-id="8a6c0-682">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-682">Name</span></span>|<span data-ttu-id="8a6c0-683">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-683">Type</span></span>|<span data-ttu-id="8a6c0-684">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-684">Attributes</span></span>|<span data-ttu-id="8a6c0-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="8a6c0-686">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-686">String</span></span>||<span data-ttu-id="8a6c0-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8a6c0-689">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-689">String</span></span>||<span data-ttu-id="8a6c0-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8a6c0-692">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-692">Object</span></span>|<span data-ttu-id="8a6c0-693">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-693">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-694">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-695">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-695">Object</span></span>|<span data-ttu-id="8a6c0-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-696">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-697">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8a6c0-698">Booliano</span><span class="sxs-lookup"><span data-stu-id="8a6c0-698">Boolean</span></span>|<span data-ttu-id="8a6c0-699">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-699">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-700">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-701">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-701">function</span></span>|<span data-ttu-id="8a6c0-702">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-702">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-703">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a6c0-704">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8a6c0-705">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a6c0-706">Erros</span><span class="sxs-lookup"><span data-stu-id="8a6c0-706">Errors</span></span>

|<span data-ttu-id="8a6c0-707">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-707">Error code</span></span>|<span data-ttu-id="8a6c0-708">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8a6c0-709">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8a6c0-710">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8a6c0-711">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-712">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-712">Requirements</span></span>

|<span data-ttu-id="8a6c0-713">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-713">Requirement</span></span>|<span data-ttu-id="8a6c0-714">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-715">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-716">1.1</span><span class="sxs-lookup"><span data-stu-id="8a6c0-716">1.1</span></span>|
|[<span data-ttu-id="8a6c0-717">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-719">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-720">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a6c0-721">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-721">Examples</span></span>

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

<span data-ttu-id="8a6c0-722">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="8a6c0-723">addFileAttachmentFromBase64Async (base64File, attachmentName, [Opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-724">Adiciona um arquivo a partir do base64 codificação para uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8a6c0-725">O `addFileAttachmentFromBase64Async` método carrega o arquivo a partir da codificação base64 e o anexa ao item no formato redação.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="8a6c0-726">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="8a6c0-727">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-728">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-728">Parameters:</span></span>
|<span data-ttu-id="8a6c0-729">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-729">Name</span></span>|<span data-ttu-id="8a6c0-730">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-730">Type</span></span>|<span data-ttu-id="8a6c0-731">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-731">Attributes</span></span>|<span data-ttu-id="8a6c0-732">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="8a6c0-733">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-733">String</span></span>||<span data-ttu-id="8a6c0-734">O base64 codificado conteúdo de uma imagem ou um arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="8a6c0-735">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-735">String</span></span>||<span data-ttu-id="8a6c0-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8a6c0-738">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-738">Object</span></span>|<span data-ttu-id="8a6c0-739">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-739">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-740">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-741">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-741">Object</span></span>|<span data-ttu-id="8a6c0-742">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-742">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-743">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8a6c0-744">Booliano</span><span class="sxs-lookup"><span data-stu-id="8a6c0-744">Boolean</span></span>|<span data-ttu-id="8a6c0-745">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-745">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-746">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-747">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-747">function</span></span>|<span data-ttu-id="8a6c0-748">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-748">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-749">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a6c0-750">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8a6c0-751">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a6c0-752">Erros</span><span class="sxs-lookup"><span data-stu-id="8a6c0-752">Errors</span></span>

|<span data-ttu-id="8a6c0-753">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-753">Error code</span></span>|<span data-ttu-id="8a6c0-754">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8a6c0-755">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8a6c0-756">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8a6c0-757">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-758">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-758">Requirements</span></span>

|<span data-ttu-id="8a6c0-759">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-759">Requirement</span></span>|<span data-ttu-id="8a6c0-760">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-761">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-762">Visualização</span><span class="sxs-lookup"><span data-stu-id="8a6c0-762">Preview</span></span>|
|[<span data-ttu-id="8a6c0-763">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-765">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-766">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a6c0-767">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-767">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8a6c0-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-769">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8a6c0-770">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8a6c0-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-771">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-771">Parameters:</span></span>

| <span data-ttu-id="8a6c0-772">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-772">Name</span></span> | <span data-ttu-id="8a6c0-773">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-773">Type</span></span> | <span data-ttu-id="8a6c0-774">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-774">Attributes</span></span> | <span data-ttu-id="8a6c0-775">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8a6c0-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8a6c0-777">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8a6c0-778">Função</span><span class="sxs-lookup"><span data-stu-id="8a6c0-778">Function</span></span> || <span data-ttu-id="8a6c0-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8a6c0-782">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-782">Object</span></span> | <span data-ttu-id="8a6c0-783">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-783">&lt;optional&gt;</span></span> | <span data-ttu-id="8a6c0-784">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8a6c0-785">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-785">Object</span></span> | <span data-ttu-id="8a6c0-786">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-786">&lt;optional&gt;</span></span> | <span data-ttu-id="8a6c0-787">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8a6c0-788">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-788">function</span></span>| <span data-ttu-id="8a6c0-789">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-789">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-790">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-791">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-791">Requirements</span></span>

|<span data-ttu-id="8a6c0-792">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-792">Requirement</span></span>| <span data-ttu-id="8a6c0-793">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-794">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a6c0-795">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-795">1.7</span></span> |
|[<span data-ttu-id="8a6c0-796">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a6c0-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-797">ReadItem</span></span> |
|[<span data-ttu-id="8a6c0-798">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a6c0-799">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8a6c0-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-801">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8a6c0-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8a6c0-805">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8a6c0-806">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-807">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-807">Parameters:</span></span>

|<span data-ttu-id="8a6c0-808">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-808">Name</span></span>|<span data-ttu-id="8a6c0-809">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-809">Type</span></span>|<span data-ttu-id="8a6c0-810">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-810">Attributes</span></span>|<span data-ttu-id="8a6c0-811">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="8a6c0-812">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-812">String</span></span>||<span data-ttu-id="8a6c0-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8a6c0-815">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-815">String</span></span>||<span data-ttu-id="8a6c0-p141">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8a6c0-818">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-818">Object</span></span>|<span data-ttu-id="8a6c0-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-819">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-820">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-821">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-821">Object</span></span>|<span data-ttu-id="8a6c0-822">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-822">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-823">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-824">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-824">function</span></span>|<span data-ttu-id="8a6c0-825">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-825">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-826">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a6c0-827">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8a6c0-828">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a6c0-829">Erros</span><span class="sxs-lookup"><span data-stu-id="8a6c0-829">Errors</span></span>

|<span data-ttu-id="8a6c0-830">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-830">Error code</span></span>|<span data-ttu-id="8a6c0-831">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8a6c0-832">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-833">Requirements</span></span>

|<span data-ttu-id="8a6c0-834">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-834">Requirement</span></span>|<span data-ttu-id="8a6c0-835">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-836">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-837">1.1</span><span class="sxs-lookup"><span data-stu-id="8a6c0-837">1.1</span></span>|
|[<span data-ttu-id="8a6c0-838">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-840">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-841">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-842">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-842">Example</span></span>

<span data-ttu-id="8a6c0-843">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8a6c0-844">close()</span><span class="sxs-lookup"><span data-stu-id="8a6c0-844">close()</span></span>

<span data-ttu-id="8a6c0-845">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8a6c0-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-848">No Outlook na web, se o item é um compromisso e ele tenha sido salva anteriormente usando `saveAsync`, o usuário é solicitado a salvar, descartar ou Cancelar, mesmo se nenhuma alteração ocorreram desde que o item foi último salvo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8a6c0-849">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-850">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-850">Requirements</span></span>

|<span data-ttu-id="8a6c0-851">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-851">Requirement</span></span>|<span data-ttu-id="8a6c0-852">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-853">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-854">1.3</span><span class="sxs-lookup"><span data-stu-id="8a6c0-854">1.3</span></span>|
|[<span data-ttu-id="8a6c0-855">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-856">Restrito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-856">Restricted</span></span>|
|[<span data-ttu-id="8a6c0-857">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-858">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8a6c0-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8a6c0-860">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-861">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-862">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8a6c0-863">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8a6c0-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-867">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-867">Parameters:</span></span>

|<span data-ttu-id="8a6c0-868">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-868">Name</span></span>|<span data-ttu-id="8a6c0-869">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-869">Type</span></span>|<span data-ttu-id="8a6c0-870">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-870">Attributes</span></span>|<span data-ttu-id="8a6c0-871">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8a6c0-872">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8a6c0-872">String &#124; Object</span></span>||<span data-ttu-id="8a6c0-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8a6c0-875">**OU**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-875">**OR**</span></span><br/><span data-ttu-id="8a6c0-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8a6c0-878">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-878">String</span></span>|<span data-ttu-id="8a6c0-879">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-879">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8a6c0-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8a6c0-883">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-883">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-884">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8a6c0-885">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-885">String</span></span>||<span data-ttu-id="8a6c0-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8a6c0-888">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-888">String</span></span>||<span data-ttu-id="8a6c0-889">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8a6c0-890">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-890">String</span></span>||<span data-ttu-id="8a6c0-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8a6c0-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="8a6c0-893">Boolean</span></span>||<span data-ttu-id="8a6c0-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8a6c0-896">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-896">String</span></span>||<span data-ttu-id="8a6c0-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-900">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-900">function</span></span>|<span data-ttu-id="8a6c0-901">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-901">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-902">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-903">Requirements</span></span>

|<span data-ttu-id="8a6c0-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-904">Requirement</span></span>|<span data-ttu-id="8a6c0-905">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-906">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-907">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-907">1.0</span></span>|
|[<span data-ttu-id="8a6c0-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-909">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-910">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-911">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a6c0-912">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-912">Examples</span></span>

<span data-ttu-id="8a6c0-913">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8a6c0-914">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8a6c0-915">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8a6c0-916">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8a6c0-917">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8a6c0-918">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8a6c0-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="8a6c0-920">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-921">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-922">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8a6c0-923">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8a6c0-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-927">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-927">Parameters:</span></span>

|<span data-ttu-id="8a6c0-928">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-928">Name</span></span>|<span data-ttu-id="8a6c0-929">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-929">Type</span></span>|<span data-ttu-id="8a6c0-930">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-930">Attributes</span></span>|<span data-ttu-id="8a6c0-931">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8a6c0-932">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8a6c0-932">String &#124; Object</span></span>||<span data-ttu-id="8a6c0-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8a6c0-935">**OU**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-935">**OR**</span></span><br/><span data-ttu-id="8a6c0-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8a6c0-938">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-938">String</span></span>|<span data-ttu-id="8a6c0-939">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-939">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8a6c0-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8a6c0-943">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-943">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-944">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8a6c0-945">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-945">String</span></span>||<span data-ttu-id="8a6c0-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8a6c0-948">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-948">String</span></span>||<span data-ttu-id="8a6c0-949">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8a6c0-950">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-950">String</span></span>||<span data-ttu-id="8a6c0-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8a6c0-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="8a6c0-953">Boolean</span></span>||<span data-ttu-id="8a6c0-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8a6c0-956">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-956">String</span></span>||<span data-ttu-id="8a6c0-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-960">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-960">function</span></span>|<span data-ttu-id="8a6c0-961">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-961">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-962">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-963">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-963">Requirements</span></span>

|<span data-ttu-id="8a6c0-964">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-964">Requirement</span></span>|<span data-ttu-id="8a6c0-965">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-966">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-967">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-967">1.0</span></span>|
|[<span data-ttu-id="8a6c0-968">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-969">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-970">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-971">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a6c0-972">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-972">Examples</span></span>

<span data-ttu-id="8a6c0-973">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8a6c0-974">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8a6c0-975">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8a6c0-976">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8a6c0-977">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8a6c0-978">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8a6c0-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8a6c0-980">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-981">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-982">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-982">Requirements</span></span>

|<span data-ttu-id="8a6c0-983">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-983">Requirement</span></span>|<span data-ttu-id="8a6c0-984">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-985">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-986">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-986">1.0</span></span>|
|[<span data-ttu-id="8a6c0-987">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-988">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-989">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-990">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-991">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-991">Returns:</span></span>

<span data-ttu-id="8a6c0-992">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8a6c0-993">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-993">Example</span></span>

<span data-ttu-id="8a6c0-994">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8a6c0-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8a6c0-996">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-997">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-998">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-998">Parameters:</span></span>

|<span data-ttu-id="8a6c0-999">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-999">Name</span></span>|<span data-ttu-id="8a6c0-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1000">Type</span></span>|<span data-ttu-id="8a6c0-1001">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="8a6c0-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="8a6c0-1003">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1004">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1004">Requirements</span></span>

|<span data-ttu-id="8a6c0-1005">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1005">Requirement</span></span>|<span data-ttu-id="8a6c0-1006">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1007">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1008">1.0</span></span>|
|[<span data-ttu-id="8a6c0-1009">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1010">Restrito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1010">Restricted</span></span>|
|[<span data-ttu-id="8a6c0-1011">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1012">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1013">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1013">Returns:</span></span>

<span data-ttu-id="8a6c0-1014">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8a6c0-1015">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8a6c0-1016">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8a6c0-1017">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="8a6c0-1018">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1018">Value of `entityType`</span></span>|<span data-ttu-id="8a6c0-1019">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1019">Type of objects in returned array</span></span>|<span data-ttu-id="8a6c0-1020">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="8a6c0-1021">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1021">String</span></span>|<span data-ttu-id="8a6c0-1022">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="8a6c0-1023">Contato</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1023">Contact</span></span>|<span data-ttu-id="8a6c0-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="8a6c0-1025">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1025">String</span></span>|<span data-ttu-id="8a6c0-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="8a6c0-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1027">MeetingSuggestion</span></span>|<span data-ttu-id="8a6c0-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="8a6c0-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1029">PhoneNumber</span></span>|<span data-ttu-id="8a6c0-1030">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="8a6c0-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1031">TaskSuggestion</span></span>|<span data-ttu-id="8a6c0-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="8a6c0-1033">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1033">String</span></span>|<span data-ttu-id="8a6c0-1034">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1034">**Restricted**</span></span>|

<span data-ttu-id="8a6c0-1035">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8a6c0-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8a6c0-1036">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1036">Example</span></span>

<span data-ttu-id="8a6c0-1037">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8a6c0-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8a6c0-1039">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1040">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-1041">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1042">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1042">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1043">Name</span></span>|<span data-ttu-id="8a6c0-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1044">Type</span></span>|<span data-ttu-id="8a6c0-1045">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8a6c0-1046">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1046">String</span></span>|<span data-ttu-id="8a6c0-1047">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1048">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1048">Requirements</span></span>

|<span data-ttu-id="8a6c0-1049">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1049">Requirement</span></span>|<span data-ttu-id="8a6c0-1050">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1051">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1052">1.0</span></span>|
|[<span data-ttu-id="8a6c0-1053">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1054">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1055">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1056">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1057">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1057">Returns:</span></span>

<span data-ttu-id="8a6c0-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8a6c0-1060">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8a6c0-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="8a6c0-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="8a6c0-1062">Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1063">Este método só é suportado pelo Outlook 2016 ou posterior para Windows (versões em Click-to-Run posterior à 16.0.8413.1000) e o Outlook na web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1064">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1064">Parameters:</span></span>
|<span data-ttu-id="8a6c0-1065">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1065">Name</span></span>|<span data-ttu-id="8a6c0-1066">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1066">Type</span></span>|<span data-ttu-id="8a6c0-1067">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1067">Attributes</span></span>|<span data-ttu-id="8a6c0-1068">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8a6c0-1069">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1069">Object</span></span>|<span data-ttu-id="8a6c0-1070">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1071">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1072">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1072">Object</span></span>|<span data-ttu-id="8a6c0-1073">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1074">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1075">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1075">function</span></span>|<span data-ttu-id="8a6c0-1076">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1077">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a6c0-1078">Em caso de sucesso, os dados de inicialização são fornecidos no `asyncResult.value` a propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="8a6c0-1079">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1080">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1080">Requirements</span></span>

|<span data-ttu-id="8a6c0-1081">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1081">Requirement</span></span>|<span data-ttu-id="8a6c0-1082">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1083">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1084">Visualização</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1084">Preview</span></span>|
|[<span data-ttu-id="8a6c0-1085">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1086">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1087">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1088">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-1089">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1089">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="8a6c0-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8a6c0-1091">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1092">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8a6c0-1096">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8a6c0-1097">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8a6c0-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1101">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1101">Requirements</span></span>

|<span data-ttu-id="8a6c0-1102">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1102">Requirement</span></span>|<span data-ttu-id="8a6c0-1103">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1104">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1105">1.0</span></span>|
|[<span data-ttu-id="8a6c0-1106">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1107">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1108">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1109">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1110">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1110">Returns:</span></span>

<span data-ttu-id="8a6c0-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8a6c0-1113">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a6c0-1114">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a6c0-1115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1115">Example</span></span>

<span data-ttu-id="8a6c0-1116">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8a6c0-1117">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8a6c0-1118">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1119">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-1120">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8a6c0-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1123">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1123">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1124">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1124">Name</span></span>|<span data-ttu-id="8a6c0-1125">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1125">Type</span></span>|<span data-ttu-id="8a6c0-1126">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8a6c0-1127">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1127">String</span></span>|<span data-ttu-id="8a6c0-1128">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1129">Requirements</span></span>

|<span data-ttu-id="8a6c0-1130">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1130">Requirement</span></span>|<span data-ttu-id="8a6c0-1131">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1132">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1133">1.0</span></span>|
|[<span data-ttu-id="8a6c0-1134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1135">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1137">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1138">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1138">Returns:</span></span>

<span data-ttu-id="8a6c0-1139">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8a6c0-1140">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a6c0-1141">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="8a6c0-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a6c0-1142">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8a6c0-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8a6c0-1144">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8a6c0-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1147">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1147">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1148">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1148">Name</span></span>|<span data-ttu-id="8a6c0-1149">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1149">Type</span></span>|<span data-ttu-id="8a6c0-1150">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1150">Attributes</span></span>|<span data-ttu-id="8a6c0-1151">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="8a6c0-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8a6c0-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="8a6c0-1156">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1156">Object</span></span>|<span data-ttu-id="8a6c0-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1158">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1159">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1159">Object</span></span>|<span data-ttu-id="8a6c0-1160">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1161">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1162">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1162">function</span></span>||<span data-ttu-id="8a6c0-1163">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a6c0-1164">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8a6c0-1165">Para acessar a propriedade source que a seleção é proveniente, chamar `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1166">Requirements</span></span>

|<span data-ttu-id="8a6c0-1167">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1167">Requirement</span></span>|<span data-ttu-id="8a6c0-1168">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1170">1.2</span></span>|
|[<span data-ttu-id="8a6c0-1171">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-1173">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1174">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1175">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1175">Returns:</span></span>

<span data-ttu-id="8a6c0-1176">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8a6c0-1177">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a6c0-1178">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a6c0-1179">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8a6c0-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8a6c0-p168">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1183">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1184">Requirements</span></span>

|<span data-ttu-id="8a6c0-1185">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1185">Requirement</span></span>|<span data-ttu-id="8a6c0-1186">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1188">1.6</span></span>|
|[<span data-ttu-id="8a6c0-1189">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1190">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1192">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1193">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1193">Returns:</span></span>

<span data-ttu-id="8a6c0-1194">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8a6c0-1195">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1195">Example</span></span>

<span data-ttu-id="8a6c0-1196">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8a6c0-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8a6c0-p169">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1200">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8a6c0-p170">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8a6c0-1204">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8a6c0-1205">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8a6c0-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1209">Requirements</span></span>

|<span data-ttu-id="8a6c0-1210">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1210">Requirement</span></span>|<span data-ttu-id="8a6c0-1211">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1213">1.6</span></span>|
|[<span data-ttu-id="8a6c0-1214">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1215">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1217">Leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a6c0-1218">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1218">Returns:</span></span>

<span data-ttu-id="8a6c0-p172">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8a6c0-1221">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1221">Example</span></span>

<span data-ttu-id="8a6c0-1222">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="8a6c0-1223">getSharedPropertiesAsync ([Opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="8a6c0-1224">Obtém as propriedades do compromisso selecionado ou da mensagem em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1225">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1225">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1226">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1226">Name</span></span>|<span data-ttu-id="8a6c0-1227">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1227">Type</span></span>|<span data-ttu-id="8a6c0-1228">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1228">Attributes</span></span>|<span data-ttu-id="8a6c0-1229">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8a6c0-1230">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1230">Object</span></span>|<span data-ttu-id="8a6c0-1231">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1232">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1233">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1233">Object</span></span>|<span data-ttu-id="8a6c0-1234">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1235">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1236">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1236">function</span></span>||<span data-ttu-id="8a6c0-1237">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a6c0-1238">As propriedades compartilhadas são fornecidas como uma [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) do objeto no `asyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8a6c0-1239">Este objeto pode ser usado para obter as propriedades do item compartilhado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1240">Requirements</span></span>

|<span data-ttu-id="8a6c0-1241">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1241">Requirement</span></span>|<span data-ttu-id="8a6c0-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1244">Visualização</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1244">Preview</span></span>|
|[<span data-ttu-id="8a6c0-1245">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1246">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1247">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1248">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-1249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8a6c0-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8a6c0-1251">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8a6c0-p174">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1255">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1255">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1256">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1256">Name</span></span>|<span data-ttu-id="8a6c0-1257">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1257">Type</span></span>|<span data-ttu-id="8a6c0-1258">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1258">Attributes</span></span>|<span data-ttu-id="8a6c0-1259">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="8a6c0-1260">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1260">function</span></span>||<span data-ttu-id="8a6c0-1261">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a6c0-1262">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8a6c0-1263">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="8a6c0-1264">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1264">Object</span></span>|<span data-ttu-id="8a6c0-1265">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1266">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8a6c0-1267">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1268">Requirements</span></span>

|<span data-ttu-id="8a6c0-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1269">Requirement</span></span>|<span data-ttu-id="8a6c0-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1272">1.0</span></span>|
|[<span data-ttu-id="8a6c0-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1274">ReadItem</span></span>|
|[<span data-ttu-id="8a6c0-1275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1276">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-1277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1277">Example</span></span>

<span data-ttu-id="8a6c0-p177">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8a6c0-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-1282">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8a6c0-p178">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1287">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1287">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1288">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1288">Name</span></span>|<span data-ttu-id="8a6c0-1289">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1289">Type</span></span>|<span data-ttu-id="8a6c0-1290">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1290">Attributes</span></span>|<span data-ttu-id="8a6c0-1291">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8a6c0-1292">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1292">String</span></span>||<span data-ttu-id="8a6c0-p179">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="8a6c0-1295">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1295">Object</span></span>|<span data-ttu-id="8a6c0-1296">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1297">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1298">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1298">Object</span></span>|<span data-ttu-id="8a6c0-1299">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1300">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1301">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1301">function</span></span>|<span data-ttu-id="8a6c0-1302">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1303">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a6c0-1304">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a6c0-1305">Erros</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1305">Errors</span></span>

|<span data-ttu-id="8a6c0-1306">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1306">Error code</span></span>|<span data-ttu-id="8a6c0-1307">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="8a6c0-1308">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1309">Requirements</span></span>

|<span data-ttu-id="8a6c0-1310">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1310">Requirement</span></span>|<span data-ttu-id="8a6c0-1311">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1313">1.1</span></span>|
|[<span data-ttu-id="8a6c0-1314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-1316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1317">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-1318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1318">Example</span></span>

<span data-ttu-id="8a6c0-1319">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8a6c0-1320">removeHandlerAsync (eventType, manipulador, [Opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8a6c0-1321">Remove um manipulador de eventos para um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="8a6c0-1322">No momento os tipos de evento aceitos são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1323">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1323">Parameters:</span></span>

| <span data-ttu-id="8a6c0-1324">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1324">Name</span></span> | <span data-ttu-id="8a6c0-1325">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1325">Type</span></span> | <span data-ttu-id="8a6c0-1326">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1326">Attributes</span></span> | <span data-ttu-id="8a6c0-1327">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8a6c0-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8a6c0-1329">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8a6c0-1330">Função</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1330">Function</span></span> || <span data-ttu-id="8a6c0-p180">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8a6c0-1334">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1334">Object</span></span> | <span data-ttu-id="8a6c0-1335">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="8a6c0-1336">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8a6c0-1337">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1337">Object</span></span> | <span data-ttu-id="8a6c0-1338">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="8a6c0-1339">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8a6c0-1340">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1340">function</span></span>| <span data-ttu-id="8a6c0-1341">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1342">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1343">Requirements</span></span>

|<span data-ttu-id="8a6c0-1344">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1344">Requirement</span></span>| <span data-ttu-id="8a6c0-1345">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a6c0-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1347">1.7</span></span> |
|[<span data-ttu-id="8a6c0-1348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a6c0-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1349">ReadItem</span></span> |
|[<span data-ttu-id="8a6c0-1350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a6c0-1351">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8a6c0-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="8a6c0-1353">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="8a6c0-p181">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1357">Se seu suplemento chama `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, pode levar algum tempo antes do item realmente é sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="8a6c0-1358">Até que o item é sincronizado, usando o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8a6c0-p183">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8a6c0-1362">Os seguintes clientes tenham comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8a6c0-1363">Mac Outlook não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="8a6c0-1364">Chamar `saveAsync` em uma reunião no Outlook Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8a6c0-1365">Outlook na web sempre envia um convite ou atualizar quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1366">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1366">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1367">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1367">Name</span></span>|<span data-ttu-id="8a6c0-1368">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1368">Type</span></span>|<span data-ttu-id="8a6c0-1369">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1369">Attributes</span></span>|<span data-ttu-id="8a6c0-1370">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8a6c0-1371">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1371">Object</span></span>|<span data-ttu-id="8a6c0-1372">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1373">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1374">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1374">Object</span></span>|<span data-ttu-id="8a6c0-1375">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1376">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1377">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1377">function</span></span>||<span data-ttu-id="8a6c0-1378">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a6c0-1379">Em caso de sucesso, o identificador do item é fornecido no `asyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1380">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1380">Requirements</span></span>

|<span data-ttu-id="8a6c0-1381">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1381">Requirement</span></span>|<span data-ttu-id="8a6c0-1382">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1383">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1384">1.3</span></span>|
|[<span data-ttu-id="8a6c0-1385">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-1387">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1388">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a6c0-1389">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8a6c0-p185">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8a6c0-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8a6c0-1393">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8a6c0-p186">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a6c0-1397">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1397">Parameters:</span></span>

|<span data-ttu-id="8a6c0-1398">Nome</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1398">Name</span></span>|<span data-ttu-id="8a6c0-1399">Tipo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1399">Type</span></span>|<span data-ttu-id="8a6c0-1400">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1400">Attributes</span></span>|<span data-ttu-id="8a6c0-1401">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="8a6c0-1402">String</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1402">String</span></span>||<span data-ttu-id="8a6c0-p187">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="8a6c0-1406">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1406">Object</span></span>|<span data-ttu-id="8a6c0-1407">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1408">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8a6c0-1409">Objeto</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1409">Object</span></span>|<span data-ttu-id="8a6c0-1410">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-1411">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8a6c0-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8a6c0-1413">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="8a6c0-p188">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8a6c0-p189">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8a6c0-1418">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="8a6c0-1419">function</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1419">function</span></span>||<span data-ttu-id="8a6c0-1420">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a6c0-1421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1421">Requirements</span></span>

|<span data-ttu-id="8a6c0-1422">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1422">Requirement</span></span>|<span data-ttu-id="8a6c0-1423">Valor</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a6c0-1424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8a6c0-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1425">1.2</span></span>|
|[<span data-ttu-id="8a6c0-1426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8a6c0-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a6c0-1428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8a6c0-1429">Composição</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a6c0-1430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a6c0-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```