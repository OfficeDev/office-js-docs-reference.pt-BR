
# <a name="item"></a><span data-ttu-id="fae57-101">item</span><span class="sxs-lookup"><span data-stu-id="fae57-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="fae57-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fae57-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="fae57-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="fae57-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-105">Requirements</span></span>

|<span data-ttu-id="fae57-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-106">Requirement</span></span>| <span data-ttu-id="fae57-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-109">1.0</span></span>|
|[<span data-ttu-id="fae57-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="fae57-111">Restricted</span></span>|
|[<span data-ttu-id="fae57-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fae57-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="fae57-114">Members and methods</span></span>

| <span data-ttu-id="fae57-115">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-115">Member</span></span> | <span data-ttu-id="fae57-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fae57-117">attachments</span><span class="sxs-lookup"><span data-stu-id="fae57-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="fae57-118">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-118">Member</span></span> |
| [<span data-ttu-id="fae57-119">bcc</span><span class="sxs-lookup"><span data-stu-id="fae57-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="fae57-120">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-120">Member</span></span> |
| [<span data-ttu-id="fae57-121">body</span><span class="sxs-lookup"><span data-stu-id="fae57-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="fae57-122">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-122">Member</span></span> |
| [<span data-ttu-id="fae57-123">cc</span><span class="sxs-lookup"><span data-stu-id="fae57-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="fae57-124">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-124">Member</span></span> |
| [<span data-ttu-id="fae57-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="fae57-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="fae57-126">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-126">Member</span></span> |
| [<span data-ttu-id="fae57-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="fae57-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="fae57-128">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-128">Member</span></span> |
| [<span data-ttu-id="fae57-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="fae57-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="fae57-130">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-130">Member</span></span> |
| [<span data-ttu-id="fae57-131">end</span><span class="sxs-lookup"><span data-stu-id="fae57-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="fae57-132">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-132">Member</span></span> |
| [<span data-ttu-id="fae57-133">from</span><span class="sxs-lookup"><span data-stu-id="fae57-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="fae57-134">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-134">Member</span></span> |
| [<span data-ttu-id="fae57-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="fae57-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="fae57-136">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-136">Member</span></span> |
| [<span data-ttu-id="fae57-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="fae57-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="fae57-138">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-138">Member</span></span> |
| [<span data-ttu-id="fae57-139">itemId</span><span class="sxs-lookup"><span data-stu-id="fae57-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="fae57-140">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-140">Member</span></span> |
| [<span data-ttu-id="fae57-141">itemType</span><span class="sxs-lookup"><span data-stu-id="fae57-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="fae57-142">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-142">Member</span></span> |
| [<span data-ttu-id="fae57-143">location</span><span class="sxs-lookup"><span data-stu-id="fae57-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="fae57-144">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-144">Member</span></span> |
| [<span data-ttu-id="fae57-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="fae57-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="fae57-146">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-146">Member</span></span> |
| [<span data-ttu-id="fae57-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="fae57-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="fae57-148">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-148">Member</span></span> |
| [<span data-ttu-id="fae57-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="fae57-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="fae57-150">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-150">Member</span></span> |
| [<span data-ttu-id="fae57-151">organizer</span><span class="sxs-lookup"><span data-stu-id="fae57-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="fae57-152">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-152">Member</span></span> |
| [<span data-ttu-id="fae57-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="fae57-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="fae57-154">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-154">Member</span></span> |
| [<span data-ttu-id="fae57-155">sender</span><span class="sxs-lookup"><span data-stu-id="fae57-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="fae57-156">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-156">Member</span></span> |
| [<span data-ttu-id="fae57-157">start</span><span class="sxs-lookup"><span data-stu-id="fae57-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="fae57-158">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-158">Member</span></span> |
| [<span data-ttu-id="fae57-159">subject</span><span class="sxs-lookup"><span data-stu-id="fae57-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="fae57-160">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-160">Member</span></span> |
| [<span data-ttu-id="fae57-161">to</span><span class="sxs-lookup"><span data-stu-id="fae57-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="fae57-162">Membro</span><span class="sxs-lookup"><span data-stu-id="fae57-162">Member</span></span> |
| [<span data-ttu-id="fae57-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="fae57-164">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-164">Method</span></span> |
| [<span data-ttu-id="fae57-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="fae57-166">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-166">Method</span></span> |
| [<span data-ttu-id="fae57-167">close</span><span class="sxs-lookup"><span data-stu-id="fae57-167">close</span></span>](#close) | <span data-ttu-id="fae57-168">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-168">Method</span></span> |
| [<span data-ttu-id="fae57-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="fae57-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="fae57-170">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-170">Method</span></span> |
| [<span data-ttu-id="fae57-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="fae57-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="fae57-172">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-172">Method</span></span> |
| [<span data-ttu-id="fae57-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="fae57-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="fae57-174">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-174">Method</span></span> |
| [<span data-ttu-id="fae57-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="fae57-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="fae57-176">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-176">Method</span></span> |
| [<span data-ttu-id="fae57-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="fae57-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="fae57-178">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-178">Method</span></span> |
| [<span data-ttu-id="fae57-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="fae57-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="fae57-180">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-180">Method</span></span> |
| [<span data-ttu-id="fae57-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="fae57-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="fae57-182">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-182">Method</span></span> |
| [<span data-ttu-id="fae57-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="fae57-184">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-184">Method</span></span> |
| [<span data-ttu-id="fae57-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="fae57-186">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-186">Method</span></span> |
| [<span data-ttu-id="fae57-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="fae57-188">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-188">Method</span></span> |
| [<span data-ttu-id="fae57-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="fae57-190">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-190">Method</span></span> |
| [<span data-ttu-id="fae57-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fae57-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="fae57-192">Método</span><span class="sxs-lookup"><span data-stu-id="fae57-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="fae57-193">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-193">Example</span></span>

<span data-ttu-id="fae57-194">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="fae57-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="fae57-195">Membros</span><span class="sxs-lookup"><span data-stu-id="fae57-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="fae57-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fae57-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="fae57-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-199">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="fae57-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fae57-200">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="fae57-200">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-201">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-201">Type:</span></span>

*   <span data-ttu-id="fae57-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fae57-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-203">Requirements</span></span>

|<span data-ttu-id="fae57-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-204">Requirement</span></span>| <span data-ttu-id="fae57-205">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-207">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-207">1.0</span></span>|
|[<span data-ttu-id="fae57-208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-209">ReadItem</span></span>|
|[<span data-ttu-id="fae57-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-211">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-212">Example</span></span>

<span data-ttu-id="fae57-213">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fae57-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fae57-215">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fae57-216">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fae57-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-217">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-217">Type:</span></span>

*   [<span data-ttu-id="fae57-218">Destinatários</span><span class="sxs-lookup"><span data-stu-id="fae57-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fae57-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-219">Requirements</span></span>

|<span data-ttu-id="fae57-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-220">Requirement</span></span>| <span data-ttu-id="fae57-221">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-223">1.1</span><span class="sxs-lookup"><span data-stu-id="fae57-223">1.1</span></span>|
|[<span data-ttu-id="fae57-224">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-225">ReadItem</span></span>|
|[<span data-ttu-id="fae57-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-227">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-228">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="fae57-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="fae57-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="fae57-230">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="fae57-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-231">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-231">Type:</span></span>

*   [<span data-ttu-id="fae57-232">Corpo</span><span class="sxs-lookup"><span data-stu-id="fae57-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="fae57-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-233">Requirements</span></span>

|<span data-ttu-id="fae57-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-234">Requirement</span></span>| <span data-ttu-id="fae57-235">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-236">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-237">1.1</span><span class="sxs-lookup"><span data-stu-id="fae57-237">1.1</span></span>|
|[<span data-ttu-id="fae57-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-239">ReadItem</span></span>|
|[<span data-ttu-id="fae57-240">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-241">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fae57-242">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fae57-243">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fae57-244">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-245">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-245">Read mode</span></span>

<span data-ttu-id="fae57-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="fae57-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-248">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-248">Compose mode</span></span>

<span data-ttu-id="fae57-249">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-250">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-250">Type:</span></span>

*   <span data-ttu-id="fae57-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-252">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-252">Requirements</span></span>

|<span data-ttu-id="fae57-253">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-253">Requirement</span></span>| <span data-ttu-id="fae57-254">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-255">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-255">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-256">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-256">1.0</span></span>|
|[<span data-ttu-id="fae57-257">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-258">ReadItem</span></span>|
|[<span data-ttu-id="fae57-259">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-260">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-261">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fae57-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fae57-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="fae57-263">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="fae57-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fae57-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="fae57-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fae57-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="fae57-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-268">Type:</span></span>

*   <span data-ttu-id="fae57-269">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-270">Requirements</span></span>

|<span data-ttu-id="fae57-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-271">Requirement</span></span>| <span data-ttu-id="fae57-272">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-274">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-274">1.0</span></span>|
|[<span data-ttu-id="fae57-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-276">ReadItem</span></span>|
|[<span data-ttu-id="fae57-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-278">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="fae57-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fae57-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="fae57-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-282">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-282">Type:</span></span>

*   <span data-ttu-id="fae57-283">Data</span><span class="sxs-lookup"><span data-stu-id="fae57-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-284">Requirements</span></span>

|<span data-ttu-id="fae57-285">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-285">Requirement</span></span>| <span data-ttu-id="fae57-286">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-287">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-287">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-288">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-288">1.0</span></span>|
|[<span data-ttu-id="fae57-289">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-290">ReadItem</span></span>|
|[<span data-ttu-id="fae57-291">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-292">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-293">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fae57-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fae57-294">dateTimeModified :Date</span></span>

<span data-ttu-id="fae57-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-297">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-297">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-298">Type:</span></span>

*   <span data-ttu-id="fae57-299">Data</span><span class="sxs-lookup"><span data-stu-id="fae57-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-300">Requirements</span></span>

|<span data-ttu-id="fae57-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-301">Requirement</span></span>| <span data-ttu-id="fae57-302">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-304">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-304">1.0</span></span>|
|[<span data-ttu-id="fae57-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-306">ReadItem</span></span>|
|[<span data-ttu-id="fae57-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-308">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="fae57-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fae57-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="fae57-311">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="fae57-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fae57-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="fae57-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-314">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-314">Read mode</span></span>

<span data-ttu-id="fae57-315">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="fae57-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-316">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-316">Compose mode</span></span>

<span data-ttu-id="fae57-317">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fae57-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fae57-318">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="fae57-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-319">Type:</span></span>

*   <span data-ttu-id="fae57-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fae57-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-321">Requirements</span></span>

|<span data-ttu-id="fae57-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-322">Requirement</span></span>| <span data-ttu-id="fae57-323">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-325">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-325">1.0</span></span>|
|[<span data-ttu-id="fae57-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-327">ReadItem</span></span>|
|[<span data-ttu-id="fae57-328">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-329">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-330">Example</span></span>

<span data-ttu-id="fae57-331">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fae57-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fae57-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fae57-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fae57-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="fae57-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="fae57-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-337">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fae57-337">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-338">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-338">Type:</span></span>

*   [<span data-ttu-id="fae57-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fae57-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fae57-340">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-340">Requirements</span></span>

|<span data-ttu-id="fae57-341">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-341">Requirement</span></span>| <span data-ttu-id="fae57-342">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-343">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-343">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-344">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-344">1.0</span></span>|
|[<span data-ttu-id="fae57-345">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-346">ReadItem</span></span>|
|[<span data-ttu-id="fae57-347">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-348">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="fae57-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fae57-349">internetMessageId :String</span></span>

<span data-ttu-id="fae57-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-352">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-352">Type:</span></span>

*   <span data-ttu-id="fae57-353">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-354">Requirements</span></span>

|<span data-ttu-id="fae57-355">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-355">Requirement</span></span>| <span data-ttu-id="fae57-356">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-357">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-358">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-358">1.0</span></span>|
|[<span data-ttu-id="fae57-359">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-360">ReadItem</span></span>|
|[<span data-ttu-id="fae57-361">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-362">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-363">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fae57-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fae57-364">itemClass :String</span></span>

<span data-ttu-id="fae57-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fae57-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="fae57-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-369">Type</span></span> | <span data-ttu-id="fae57-370">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-370">Description</span></span> | <span data-ttu-id="fae57-371">classe de item</span><span class="sxs-lookup"><span data-stu-id="fae57-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="fae57-372">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="fae57-372">Appointment items</span></span> | <span data-ttu-id="fae57-373">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="fae57-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="fae57-374">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="fae57-374">Message items</span></span> | <span data-ttu-id="fae57-375">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="fae57-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="fae57-376">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="fae57-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-377">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-377">Type:</span></span>

*   <span data-ttu-id="fae57-378">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-379">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-379">Requirements</span></span>

|<span data-ttu-id="fae57-380">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-380">Requirement</span></span>| <span data-ttu-id="fae57-381">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-382">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-382">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-383">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-383">1.0</span></span>|
|[<span data-ttu-id="fae57-384">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-385">ReadItem</span></span>|
|[<span data-ttu-id="fae57-386">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-387">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-388">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fae57-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fae57-389">(nullable) itemId :String</span></span>

<span data-ttu-id="fae57-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-392">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="fae57-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fae57-393">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fae57-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fae57-394">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fae57-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fae57-395">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="fae57-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="fae57-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-398">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-398">Type:</span></span>

*   <span data-ttu-id="fae57-399">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-400">Requirements</span></span>

|<span data-ttu-id="fae57-401">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-401">Requirement</span></span>| <span data-ttu-id="fae57-402">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-403">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-404">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-404">1.0</span></span>|
|[<span data-ttu-id="fae57-405">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-406">ReadItem</span></span>|
|[<span data-ttu-id="fae57-407">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-408">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-409">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-409">Example</span></span>

<span data-ttu-id="fae57-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fae57-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="fae57-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fae57-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fae57-413">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="fae57-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fae57-414">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-415">Type:</span></span>

*   [<span data-ttu-id="fae57-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fae57-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fae57-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-417">Requirements</span></span>

|<span data-ttu-id="fae57-418">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-418">Requirement</span></span>| <span data-ttu-id="fae57-419">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-420">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-421">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-421">1.0</span></span>|
|[<span data-ttu-id="fae57-422">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-423">ReadItem</span></span>|
|[<span data-ttu-id="fae57-424">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-425">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-426">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="fae57-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="fae57-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="fae57-428">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-429">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-429">Read mode</span></span>

<span data-ttu-id="fae57-430">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-431">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-431">Compose mode</span></span>

<span data-ttu-id="fae57-432">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-433">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-433">Type:</span></span>

*   <span data-ttu-id="fae57-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="fae57-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-435">Requirements</span></span>

|<span data-ttu-id="fae57-436">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-436">Requirement</span></span>| <span data-ttu-id="fae57-437">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-438">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-438">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-439">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-439">1.0</span></span>|
|[<span data-ttu-id="fae57-440">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-441">ReadItem</span></span>|
|[<span data-ttu-id="fae57-442">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-443">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-444">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fae57-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fae57-445">normalizedSubject :String</span></span>

<span data-ttu-id="fae57-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fae57-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="fae57-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-450">Type:</span></span>

*   <span data-ttu-id="fae57-451">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-452">Requirements</span></span>

|<span data-ttu-id="fae57-453">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-453">Requirement</span></span>| <span data-ttu-id="fae57-454">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-455">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-456">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-456">1.0</span></span>|
|[<span data-ttu-id="fae57-457">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-458">ReadItem</span></span>|
|[<span data-ttu-id="fae57-459">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-460">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-461">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="fae57-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="fae57-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="fae57-463">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="fae57-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-464">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-464">Type:</span></span>

*   [<span data-ttu-id="fae57-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="fae57-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="fae57-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-466">Requirements</span></span>

|<span data-ttu-id="fae57-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-467">Requirement</span></span>| <span data-ttu-id="fae57-468">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-469">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-470">1.3</span><span class="sxs-lookup"><span data-stu-id="fae57-470">1.3</span></span>|
|[<span data-ttu-id="fae57-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-472">ReadItem</span></span>|
|[<span data-ttu-id="fae57-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-474">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fae57-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fae57-476">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="fae57-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fae57-477">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-478">Read mode</span></span>

<span data-ttu-id="fae57-479">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="fae57-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-480">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-480">Compose mode</span></span>

<span data-ttu-id="fae57-481">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="fae57-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-482">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-482">Type:</span></span>

*   <span data-ttu-id="fae57-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-484">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-484">Requirements</span></span>

|<span data-ttu-id="fae57-485">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-485">Requirement</span></span>| <span data-ttu-id="fae57-486">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-487">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-487">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-488">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-488">1.0</span></span>|
|[<span data-ttu-id="fae57-489">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-490">ReadItem</span></span>|
|[<span data-ttu-id="fae57-491">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-492">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-493">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fae57-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fae57-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fae57-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-497">Type:</span></span>

*   [<span data-ttu-id="fae57-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fae57-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fae57-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-499">Requirements</span></span>

|<span data-ttu-id="fae57-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-500">Requirement</span></span>| <span data-ttu-id="fae57-501">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-503">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-503">1.0</span></span>|
|[<span data-ttu-id="fae57-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-505">ReadItem</span></span>|
|[<span data-ttu-id="fae57-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-507">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-508">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fae57-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fae57-510">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="fae57-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fae57-511">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-512">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-512">Read mode</span></span>

<span data-ttu-id="fae57-513">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="fae57-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-514">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-514">Compose mode</span></span>

<span data-ttu-id="fae57-515">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="fae57-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-516">Type:</span></span>

*   <span data-ttu-id="fae57-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-518">Requirements</span></span>

|<span data-ttu-id="fae57-519">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-519">Requirement</span></span>| <span data-ttu-id="fae57-520">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-521">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-522">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-522">1.0</span></span>|
|[<span data-ttu-id="fae57-523">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-524">ReadItem</span></span>|
|[<span data-ttu-id="fae57-525">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-526">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-527">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fae57-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fae57-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fae57-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fae57-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fae57-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="fae57-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-533">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `sender` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fae57-533">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-534">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-534">Type:</span></span>

*   [<span data-ttu-id="fae57-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fae57-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fae57-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-536">Requirements</span></span>

|<span data-ttu-id="fae57-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-537">Requirement</span></span>| <span data-ttu-id="fae57-538">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-539">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-540">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-540">1.0</span></span>|
|[<span data-ttu-id="fae57-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-542">ReadItem</span></span>|
|[<span data-ttu-id="fae57-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-544">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="fae57-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fae57-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="fae57-547">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="fae57-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fae57-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="fae57-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-550">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-550">Read mode</span></span>

<span data-ttu-id="fae57-551">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="fae57-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-552">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-552">Compose mode</span></span>

<span data-ttu-id="fae57-553">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fae57-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fae57-554">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="fae57-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-555">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-555">Type:</span></span>

*   <span data-ttu-id="fae57-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fae57-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-557">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-557">Requirements</span></span>

|<span data-ttu-id="fae57-558">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-558">Requirement</span></span>| <span data-ttu-id="fae57-559">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-560">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-560">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-561">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-561">1.0</span></span>|
|[<span data-ttu-id="fae57-562">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-563">ReadItem</span></span>|
|[<span data-ttu-id="fae57-564">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-565">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-566">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-566">Example</span></span>

<span data-ttu-id="fae57-567">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fae57-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="fae57-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fae57-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="fae57-569">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="fae57-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fae57-570">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="fae57-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-571">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-571">Read mode</span></span>

<span data-ttu-id="fae57-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="fae57-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="fae57-574">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-574">Compose mode</span></span>

<span data-ttu-id="fae57-575">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="fae57-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fae57-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-576">Type:</span></span>

*   <span data-ttu-id="fae57-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fae57-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-578">Requirements</span></span>

|<span data-ttu-id="fae57-579">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-579">Requirement</span></span>| <span data-ttu-id="fae57-580">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-581">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-582">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-582">1.0</span></span>|
|[<span data-ttu-id="fae57-583">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-584">ReadItem</span></span>|
|[<span data-ttu-id="fae57-585">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-586">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fae57-587">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fae57-588">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fae57-589">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fae57-590">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-590">Read mode</span></span>

<span data-ttu-id="fae57-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="fae57-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fae57-593">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fae57-593">Compose mode</span></span>

<span data-ttu-id="fae57-594">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-594">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fae57-595">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fae57-595">Type:</span></span>

*   <span data-ttu-id="fae57-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fae57-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-597">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-597">Requirements</span></span>

|<span data-ttu-id="fae57-598">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-598">Requirement</span></span>| <span data-ttu-id="fae57-599">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-600">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-600">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-601">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-601">1.0</span></span>|
|[<span data-ttu-id="fae57-602">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-603">ReadItem</span></span>|
|[<span data-ttu-id="fae57-604">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-605">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-606">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="fae57-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="fae57-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fae57-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fae57-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fae57-609">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="fae57-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fae57-610">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="fae57-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fae57-611">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="fae57-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-612">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-612">Parameters:</span></span>

|<span data-ttu-id="fae57-613">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-613">Name</span></span>| <span data-ttu-id="fae57-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-614">Type</span></span>| <span data-ttu-id="fae57-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-615">Attributes</span></span>| <span data-ttu-id="fae57-616">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="fae57-617">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-617">String</span></span>||<span data-ttu-id="fae57-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fae57-620">String</span><span class="sxs-lookup"><span data-stu-id="fae57-620">String</span></span>||<span data-ttu-id="fae57-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fae57-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-623">Object</span></span>| <span data-ttu-id="fae57-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-624">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="fae57-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-626">Object</span></span> | <span data-ttu-id="fae57-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-627">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="fae57-629">Booliano</span><span class="sxs-lookup"><span data-stu-id="fae57-629">Boolean</span></span> | <span data-ttu-id="fae57-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-630">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-631">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="fae57-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="fae57-632">function</span><span class="sxs-lookup"><span data-stu-id="fae57-632">function</span></span>| <span data-ttu-id="fae57-633">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-633">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-634">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fae57-635">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fae57-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fae57-636">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="fae57-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fae57-637">Erros</span><span class="sxs-lookup"><span data-stu-id="fae57-637">Errors</span></span>

| <span data-ttu-id="fae57-638">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fae57-638">Error code</span></span> | <span data-ttu-id="fae57-639">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="fae57-640">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="fae57-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="fae57-641">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="fae57-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fae57-642">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="fae57-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-643">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-643">Requirements</span></span>

|<span data-ttu-id="fae57-644">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-644">Requirement</span></span>| <span data-ttu-id="fae57-645">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-646">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-646">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-647">1.1</span><span class="sxs-lookup"><span data-stu-id="fae57-647">1.1</span></span>|
|[<span data-ttu-id="fae57-648">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-650">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-651">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fae57-652">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fae57-652">Examples</span></span>

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

<span data-ttu-id="fae57-653">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fae57-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fae57-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fae57-655">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fae57-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="fae57-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fae57-659">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="fae57-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fae57-660">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="fae57-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-661">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-661">Parameters:</span></span>

|<span data-ttu-id="fae57-662">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-662">Name</span></span>| <span data-ttu-id="fae57-663">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-663">Type</span></span>| <span data-ttu-id="fae57-664">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-664">Attributes</span></span>| <span data-ttu-id="fae57-665">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="fae57-666">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-666">String</span></span>||<span data-ttu-id="fae57-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fae57-669">String</span><span class="sxs-lookup"><span data-stu-id="fae57-669">String</span></span>||<span data-ttu-id="fae57-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fae57-672">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-672">Object</span></span>| <span data-ttu-id="fae57-673">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-673">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-674">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fae57-675">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-675">Object</span></span>| <span data-ttu-id="fae57-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-676">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-677">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fae57-678">function</span><span class="sxs-lookup"><span data-stu-id="fae57-678">function</span></span>| <span data-ttu-id="fae57-679">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-679">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-680">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fae57-681">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fae57-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fae57-682">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="fae57-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fae57-683">Erros</span><span class="sxs-lookup"><span data-stu-id="fae57-683">Errors</span></span>

| <span data-ttu-id="fae57-684">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fae57-684">Error code</span></span> | <span data-ttu-id="fae57-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fae57-686">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="fae57-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-687">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-687">Requirements</span></span>

|<span data-ttu-id="fae57-688">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-688">Requirement</span></span>| <span data-ttu-id="fae57-689">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-690">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-690">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-691">1.1</span><span class="sxs-lookup"><span data-stu-id="fae57-691">1.1</span></span>|
|[<span data-ttu-id="fae57-692">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-694">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-695">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-696">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-696">Example</span></span>

<span data-ttu-id="fae57-697">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="fae57-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="fae57-698">close()</span><span class="sxs-lookup"><span data-stu-id="fae57-698">close()</span></span>

<span data-ttu-id="fae57-699">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="fae57-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="fae57-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="fae57-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-702">No Outlook na web, se o item é um compromisso e ele tenha sido salva anteriormente usando `saveAsync`, o usuário é solicitado a salvar, descartar ou Cancelar, mesmo se nenhuma alteração ocorreram desde que o item foi último salvo.</span><span class="sxs-lookup"><span data-stu-id="fae57-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="fae57-703">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="fae57-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-704">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-704">Requirements</span></span>

|<span data-ttu-id="fae57-705">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-705">Requirement</span></span>| <span data-ttu-id="fae57-706">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-707">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-707">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-708">1.3</span><span class="sxs-lookup"><span data-stu-id="fae57-708">1.3</span></span>|
|[<span data-ttu-id="fae57-709">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-710">Restrito</span><span class="sxs-lookup"><span data-stu-id="fae57-710">Restricted</span></span>|
|[<span data-ttu-id="fae57-711">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-712">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="fae57-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fae57-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="fae57-714">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="fae57-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-715">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fae57-716">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="fae57-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fae57-717">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fae57-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fae57-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="fae57-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-721">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-721">Parameters:</span></span>

| <span data-ttu-id="fae57-722">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-722">Name</span></span> | <span data-ttu-id="fae57-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-723">Type</span></span> | <span data-ttu-id="fae57-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-724">Attributes</span></span> | <span data-ttu-id="fae57-725">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="fae57-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fae57-726">String &#124; Object</span></span>| |<span data-ttu-id="fae57-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fae57-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fae57-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="fae57-729">**OR**</span></span><br/><span data-ttu-id="fae57-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fae57-732">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-732">String</span></span> | <span data-ttu-id="fae57-733">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-733">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fae57-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fae57-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fae57-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-737">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-738">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="fae57-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fae57-739">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-739">String</span></span> | | <span data-ttu-id="fae57-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fae57-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fae57-742">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-742">String</span></span> | | <span data-ttu-id="fae57-743">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="fae57-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fae57-744">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-744">String</span></span> | | <span data-ttu-id="fae57-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fae57-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="fae57-747">Boolean</span><span class="sxs-lookup"><span data-stu-id="fae57-747">Boolean</span></span> | | <span data-ttu-id="fae57-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="fae57-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fae57-750">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-750">String</span></span> | | <span data-ttu-id="fae57-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fae57-754">function</span><span class="sxs-lookup"><span data-stu-id="fae57-754">function</span></span> | <span data-ttu-id="fae57-755">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-755">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-756">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-757">Requirements</span></span>

|<span data-ttu-id="fae57-758">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-758">Requirement</span></span>| <span data-ttu-id="fae57-759">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-760">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-760">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-761">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-761">1.0</span></span>|
|[<span data-ttu-id="fae57-762">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-763">ReadItem</span></span>|
|[<span data-ttu-id="fae57-764">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-765">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fae57-766">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fae57-766">Examples</span></span>

<span data-ttu-id="fae57-767">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="fae57-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fae57-768">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="fae57-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fae57-769">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="fae57-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fae57-770">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="fae57-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fae57-771">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fae57-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fae57-772">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="fae57-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fae57-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="fae57-774">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="fae57-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-775">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fae57-776">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="fae57-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fae57-777">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fae57-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fae57-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="fae57-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-781">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-781">Parameters:</span></span>

| <span data-ttu-id="fae57-782">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-782">Name</span></span> | <span data-ttu-id="fae57-783">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-783">Type</span></span> | <span data-ttu-id="fae57-784">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-784">Attributes</span></span> | <span data-ttu-id="fae57-785">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="fae57-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fae57-786">String &#124; Object</span></span>| | <span data-ttu-id="fae57-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fae57-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fae57-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="fae57-789">**OR**</span></span><br/><span data-ttu-id="fae57-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fae57-792">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-792">String</span></span> | <span data-ttu-id="fae57-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-793">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fae57-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fae57-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fae57-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-797">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-798">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="fae57-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fae57-799">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-799">String</span></span> | | <span data-ttu-id="fae57-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fae57-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fae57-802">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-802">String</span></span> | | <span data-ttu-id="fae57-803">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="fae57-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fae57-804">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-804">String</span></span> | | <span data-ttu-id="fae57-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fae57-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="fae57-807">Boolean</span><span class="sxs-lookup"><span data-stu-id="fae57-807">Boolean</span></span> | | <span data-ttu-id="fae57-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="fae57-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fae57-810">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-810">String</span></span> | | <span data-ttu-id="fae57-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fae57-814">function</span><span class="sxs-lookup"><span data-stu-id="fae57-814">function</span></span> | <span data-ttu-id="fae57-815">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-815">&lt;optional&gt;</span></span> | <span data-ttu-id="fae57-816">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-817">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-817">Requirements</span></span>

|<span data-ttu-id="fae57-818">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-818">Requirement</span></span>| <span data-ttu-id="fae57-819">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-820">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-820">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-821">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-821">1.0</span></span>|
|[<span data-ttu-id="fae57-822">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-823">ReadItem</span></span>|
|[<span data-ttu-id="fae57-824">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-825">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fae57-826">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fae57-826">Examples</span></span>

<span data-ttu-id="fae57-827">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="fae57-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fae57-828">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="fae57-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fae57-829">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="fae57-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fae57-830">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="fae57-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fae57-831">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fae57-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fae57-832">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="fae57-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fae57-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="fae57-834">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fae57-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-835">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-836">Requirements</span></span>

|<span data-ttu-id="fae57-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-837">Requirement</span></span>| <span data-ttu-id="fae57-838">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-839">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-840">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-840">1.0</span></span>|
|[<span data-ttu-id="fae57-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-842">ReadItem</span></span>|
|[<span data-ttu-id="fae57-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-844">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-845">Returns:</span></span>

<span data-ttu-id="fae57-846">Tipo: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fae57-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fae57-847">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-847">Example</span></span>

<span data-ttu-id="fae57-848">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-848">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="fae57-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fae57-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fae57-850">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fae57-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-851">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-852">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-852">Parameters:</span></span>

|<span data-ttu-id="fae57-853">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-853">Name</span></span>| <span data-ttu-id="fae57-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-854">Type</span></span>| <span data-ttu-id="fae57-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="fae57-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fae57-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="fae57-857">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="fae57-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-858">Requirements</span></span>

|<span data-ttu-id="fae57-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-859">Requirement</span></span>| <span data-ttu-id="fae57-860">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-861">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-862">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-862">1.0</span></span>|
|[<span data-ttu-id="fae57-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-864">Restrito</span><span class="sxs-lookup"><span data-stu-id="fae57-864">Restricted</span></span>|
|[<span data-ttu-id="fae57-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-866">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-867">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-867">Returns:</span></span>

<span data-ttu-id="fae57-868">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="fae57-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fae57-869">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="fae57-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="fae57-870">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="fae57-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fae57-871">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="fae57-872">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="fae57-872">Value of `entityType`</span></span> | <span data-ttu-id="fae57-873">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="fae57-873">Type of objects in returned array</span></span> | <span data-ttu-id="fae57-874">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="fae57-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="fae57-875">String</span><span class="sxs-lookup"><span data-stu-id="fae57-875">String</span></span> | <span data-ttu-id="fae57-876">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fae57-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="fae57-877">Contato</span><span class="sxs-lookup"><span data-stu-id="fae57-877">Contact</span></span> | <span data-ttu-id="fae57-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fae57-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="fae57-879">String</span><span class="sxs-lookup"><span data-stu-id="fae57-879">String</span></span> | <span data-ttu-id="fae57-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fae57-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="fae57-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fae57-881">MeetingSuggestion</span></span> | <span data-ttu-id="fae57-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fae57-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="fae57-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fae57-883">PhoneNumber</span></span> | <span data-ttu-id="fae57-884">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fae57-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="fae57-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fae57-885">TaskSuggestion</span></span> | <span data-ttu-id="fae57-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fae57-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="fae57-887">String</span><span class="sxs-lookup"><span data-stu-id="fae57-887">String</span></span> | <span data-ttu-id="fae57-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fae57-888">**Restricted**</span></span> |

<span data-ttu-id="fae57-889">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fae57-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fae57-890">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-890">Example</span></span>

<span data-ttu-id="fae57-891">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fae57-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="fae57-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fae57-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fae57-893">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fae57-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-894">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fae57-895">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="fae57-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-896">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-896">Parameters:</span></span>

|<span data-ttu-id="fae57-897">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-897">Name</span></span>| <span data-ttu-id="fae57-898">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-898">Type</span></span>| <span data-ttu-id="fae57-899">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fae57-900">String</span><span class="sxs-lookup"><span data-stu-id="fae57-900">String</span></span>|<span data-ttu-id="fae57-901">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="fae57-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-902">Requirements</span></span>

|<span data-ttu-id="fae57-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-903">Requirement</span></span>| <span data-ttu-id="fae57-904">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-906">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-906">1.0</span></span>|
|[<span data-ttu-id="fae57-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-908">ReadItem</span></span>|
|[<span data-ttu-id="fae57-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-910">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-911">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-911">Returns:</span></span>

<span data-ttu-id="fae57-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="fae57-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fae57-914">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fae57-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="fae57-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fae57-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fae57-916">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fae57-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-917">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fae57-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="fae57-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fae57-921">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="fae57-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fae57-922">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fae57-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fae57-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="fae57-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fae57-926">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-926">Requirements</span></span>

|<span data-ttu-id="fae57-927">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-927">Requirement</span></span>| <span data-ttu-id="fae57-928">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-929">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-929">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-930">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-930">1.0</span></span>|
|[<span data-ttu-id="fae57-931">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-932">ReadItem</span></span>|
|[<span data-ttu-id="fae57-933">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-934">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-935">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-935">Returns:</span></span>

<span data-ttu-id="fae57-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="fae57-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fae57-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fae57-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fae57-939">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fae57-940">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-940">Example</span></span>

<span data-ttu-id="fae57-941">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="fae57-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fae57-942">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="fae57-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fae57-943">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fae57-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-944">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fae57-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fae57-945">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="fae57-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fae57-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="fae57-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-948">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-948">Parameters:</span></span>

|<span data-ttu-id="fae57-949">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-949">Name</span></span>| <span data-ttu-id="fae57-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-950">Type</span></span>| <span data-ttu-id="fae57-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fae57-952">String</span><span class="sxs-lookup"><span data-stu-id="fae57-952">String</span></span>|<span data-ttu-id="fae57-953">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="fae57-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-954">Requirements</span></span>

|<span data-ttu-id="fae57-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-955">Requirement</span></span>| <span data-ttu-id="fae57-956">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-957">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-958">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-958">1.0</span></span>|
|[<span data-ttu-id="fae57-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-960">ReadItem</span></span>|
|[<span data-ttu-id="fae57-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-962">Leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-963">Returns:</span></span>

<span data-ttu-id="fae57-964">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fae57-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fae57-965">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="fae57-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fae57-966">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="fae57-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fae57-967">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fae57-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="fae57-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fae57-969">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fae57-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="fae57-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-972">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-972">Parameters:</span></span>

|<span data-ttu-id="fae57-973">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-973">Name</span></span>| <span data-ttu-id="fae57-974">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-974">Type</span></span>| <span data-ttu-id="fae57-975">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-975">Attributes</span></span>| <span data-ttu-id="fae57-976">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="fae57-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fae57-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fae57-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="fae57-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="fae57-981">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-981">Object</span></span>| <span data-ttu-id="fae57-982">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-982">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-983">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fae57-984">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-984">Object</span></span>| <span data-ttu-id="fae57-985">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-985">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-986">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fae57-987">function</span><span class="sxs-lookup"><span data-stu-id="fae57-987">function</span></span>||<span data-ttu-id="fae57-988">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fae57-989">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="fae57-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fae57-990">Para acessar a propriedade source que a seleção é proveniente, chamar `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="fae57-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-991">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-991">Requirements</span></span>

|<span data-ttu-id="fae57-992">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-992">Requirement</span></span>| <span data-ttu-id="fae57-993">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-994">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-994">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-995">1.2</span><span class="sxs-lookup"><span data-stu-id="fae57-995">1.2</span></span>|
|[<span data-ttu-id="fae57-996">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-998">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-999">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fae57-1000">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fae57-1000">Returns:</span></span>

<span data-ttu-id="fae57-1001">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="fae57-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fae57-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fae57-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fae57-1003">String</span><span class="sxs-lookup"><span data-stu-id="fae57-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fae57-1004">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fae57-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fae57-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fae57-1006">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fae57-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fae57-p163">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="fae57-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-1010">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-1010">Parameters:</span></span>

|<span data-ttu-id="fae57-1011">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-1011">Name</span></span>| <span data-ttu-id="fae57-1012">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-1012">Type</span></span>| <span data-ttu-id="fae57-1013">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-1013">Attributes</span></span>| <span data-ttu-id="fae57-1014">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fae57-1015">function</span><span class="sxs-lookup"><span data-stu-id="fae57-1015">function</span></span>||<span data-ttu-id="fae57-1016">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fae57-1017">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fae57-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fae57-1018">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="fae57-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="fae57-1019">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1019">Object</span></span>| <span data-ttu-id="fae57-1020">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1021">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="fae57-1022">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-1023">Requirements</span></span>

|<span data-ttu-id="fae57-1024">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-1024">Requirement</span></span>| <span data-ttu-id="fae57-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-1026">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-1026">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="fae57-1027">1.0</span></span>|
|[<span data-ttu-id="fae57-1028">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fae57-1029">ReadItem</span></span>|
|[<span data-ttu-id="fae57-1030">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-1031">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fae57-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-1032">Example</span></span>

<span data-ttu-id="fae57-p166">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="fae57-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fae57-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fae57-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fae57-1037">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fae57-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fae57-p167">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="fae57-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-1042">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-1042">Parameters:</span></span>

|<span data-ttu-id="fae57-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-1043">Name</span></span>| <span data-ttu-id="fae57-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-1044">Type</span></span>| <span data-ttu-id="fae57-1045">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-1045">Attributes</span></span>| <span data-ttu-id="fae57-1046">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="fae57-1047">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fae57-1047">String</span></span>||<span data-ttu-id="fae57-p168">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae57-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="fae57-1050">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1050">Object</span></span>| <span data-ttu-id="fae57-1051">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1052">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fae57-1053">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1053">Object</span></span>| <span data-ttu-id="fae57-1054">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1055">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fae57-1056">function</span><span class="sxs-lookup"><span data-stu-id="fae57-1056">function</span></span>| <span data-ttu-id="fae57-1057">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1058">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fae57-1059">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="fae57-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fae57-1060">Erros</span><span class="sxs-lookup"><span data-stu-id="fae57-1060">Errors</span></span>

| <span data-ttu-id="fae57-1061">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fae57-1061">Error code</span></span> | <span data-ttu-id="fae57-1062">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="fae57-1063">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="fae57-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-1064">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-1064">Requirements</span></span>

|<span data-ttu-id="fae57-1065">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-1065">Requirement</span></span>| <span data-ttu-id="fae57-1066">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-1067">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-1067">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="fae57-1068">1.1</span></span>|
|[<span data-ttu-id="fae57-1069">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-1071">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-1072">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-1073">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-1073">Example</span></span>

<span data-ttu-id="fae57-1074">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="fae57-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="fae57-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fae57-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="fae57-1076">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="fae57-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="fae57-p169">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="fae57-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-1080">Se seu suplemento chama `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, pode levar algum tempo antes do item realmente é sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="fae57-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="fae57-1081">Até que o item é sincronizado, usando o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="fae57-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="fae57-p171">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="fae57-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="fae57-1085">Os seguintes clientes tenham comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="fae57-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="fae57-1086">Mac Outlook não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fae57-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="fae57-1087">Chamar `saveAsync` em uma reunião no Outlook Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="fae57-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="fae57-1088">Outlook na web sempre envia um convite ou atualizar quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fae57-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-1089">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-1089">Parameters:</span></span>

|<span data-ttu-id="fae57-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-1090">Name</span></span>| <span data-ttu-id="fae57-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-1091">Type</span></span>| <span data-ttu-id="fae57-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-1092">Attributes</span></span>| <span data-ttu-id="fae57-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="fae57-1094">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1094">Object</span></span>| <span data-ttu-id="fae57-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1096">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fae57-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1097">Object</span></span>| <span data-ttu-id="fae57-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1099">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fae57-1100">function</span><span class="sxs-lookup"><span data-stu-id="fae57-1100">function</span></span>||<span data-ttu-id="fae57-1101">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fae57-1102">Em caso de sucesso, o identificador do item é fornecido no `asyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="fae57-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fae57-1103">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-1103">Requirements</span></span>

|<span data-ttu-id="fae57-1104">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-1104">Requirement</span></span>| <span data-ttu-id="fae57-1105">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-1106">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-1106">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="fae57-1107">1.3</span></span>|
|[<span data-ttu-id="fae57-1108">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-1110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-1111">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fae57-1112">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fae57-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="fae57-p173">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="fae57-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fae57-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="fae57-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fae57-1116">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fae57-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fae57-p174">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="fae57-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fae57-1120">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fae57-1120">Parameters:</span></span>

|<span data-ttu-id="fae57-1121">Nome</span><span class="sxs-lookup"><span data-stu-id="fae57-1121">Name</span></span>| <span data-ttu-id="fae57-1122">Tipo</span><span class="sxs-lookup"><span data-stu-id="fae57-1122">Type</span></span>| <span data-ttu-id="fae57-1123">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae57-1123">Attributes</span></span>| <span data-ttu-id="fae57-1124">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae57-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fae57-1125">String</span><span class="sxs-lookup"><span data-stu-id="fae57-1125">String</span></span>||<span data-ttu-id="fae57-p175">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="fae57-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="fae57-1129">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1129">Object</span></span>| <span data-ttu-id="fae57-1130">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1131">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae57-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fae57-1132">Objeto</span><span class="sxs-lookup"><span data-stu-id="fae57-1132">Object</span></span>| <span data-ttu-id="fae57-1133">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-1134">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fae57-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="fae57-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fae57-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="fae57-1136">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fae57-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="fae57-p176">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="fae57-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fae57-p177">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="fae57-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fae57-1141">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="fae57-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="fae57-1142">function</span><span class="sxs-lookup"><span data-stu-id="fae57-1142">function</span></span>||<span data-ttu-id="fae57-1143">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fae57-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fae57-1144">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fae57-1144">Requirements</span></span>

|<span data-ttu-id="fae57-1145">Requisito</span><span class="sxs-lookup"><span data-stu-id="fae57-1145">Requirement</span></span>| <span data-ttu-id="fae57-1146">Valor</span><span class="sxs-lookup"><span data-stu-id="fae57-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="fae57-1147">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fae57-1147">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fae57-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="fae57-1148">1.2</span></span>|
|[<span data-ttu-id="fae57-1149">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fae57-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fae57-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fae57-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="fae57-1151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fae57-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fae57-1152">Composição</span><span class="sxs-lookup"><span data-stu-id="fae57-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fae57-1153">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fae57-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```