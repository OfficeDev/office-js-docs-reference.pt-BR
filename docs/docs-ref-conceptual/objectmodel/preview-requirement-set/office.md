 

# <a name="office"></a><span data-ttu-id="e33f7-101">Office</span><span class="sxs-lookup"><span data-stu-id="e33f7-101">Office</span></span>

<span data-ttu-id="e33f7-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e33f7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e33f7-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e33f7-104">Requirements</span></span>

|<span data-ttu-id="e33f7-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="e33f7-105">Requirement</span></span>| <span data-ttu-id="e33f7-106">Valor</span><span class="sxs-lookup"><span data-stu-id="e33f7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e33f7-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e33f7-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e33f7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e33f7-108">1.0</span></span>|
|[<span data-ttu-id="e33f7-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e33f7-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e33f7-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e33f7-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e33f7-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e33f7-111">Members and methods</span></span>

| <span data-ttu-id="e33f7-112">Membro</span><span class="sxs-lookup"><span data-stu-id="e33f7-112">Member</span></span> | <span data-ttu-id="e33f7-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="e33f7-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e33f7-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e33f7-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e33f7-115">Membro</span><span class="sxs-lookup"><span data-stu-id="e33f7-115">Member</span></span> |
| [<span data-ttu-id="e33f7-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e33f7-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e33f7-117">Membro</span><span class="sxs-lookup"><span data-stu-id="e33f7-117">Member</span></span> |
| [<span data-ttu-id="e33f7-118">EventType</span><span class="sxs-lookup"><span data-stu-id="e33f7-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e33f7-119">Membro</span><span class="sxs-lookup"><span data-stu-id="e33f7-119">Member</span></span> |
| [<span data-ttu-id="e33f7-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e33f7-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e33f7-121">Membro</span><span class="sxs-lookup"><span data-stu-id="e33f7-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e33f7-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="e33f7-122">Namespaces</span></span>

<span data-ttu-id="e33f7-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e33f7-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e33f7-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="e33f7-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="e33f7-125">Membros</span><span class="sxs-lookup"><span data-stu-id="e33f7-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="e33f7-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="e33f7-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="e33f7-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="e33f7-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e33f7-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e33f7-128">Type:</span></span>

*   <span data-ttu-id="e33f7-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e33f7-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e33f7-130">Properties:</span></span>

|<span data-ttu-id="e33f7-131">Nome</span><span class="sxs-lookup"><span data-stu-id="e33f7-131">Name</span></span>| <span data-ttu-id="e33f7-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="e33f7-132">Type</span></span>| <span data-ttu-id="e33f7-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="e33f7-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e33f7-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-134">String</span></span>|<span data-ttu-id="e33f7-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="e33f7-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e33f7-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-136">String</span></span>|<span data-ttu-id="e33f7-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="e33f7-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e33f7-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e33f7-138">Requirements</span></span>

|<span data-ttu-id="e33f7-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="e33f7-139">Requirement</span></span>| <span data-ttu-id="e33f7-140">Valor</span><span class="sxs-lookup"><span data-stu-id="e33f7-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="e33f7-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e33f7-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e33f7-142">1.0</span><span class="sxs-lookup"><span data-stu-id="e33f7-142">1.0</span></span>|
|[<span data-ttu-id="e33f7-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e33f7-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e33f7-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e33f7-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="e33f7-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="e33f7-145">CoercionType :String</span></span>

<span data-ttu-id="e33f7-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="e33f7-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e33f7-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e33f7-147">Type:</span></span>

*   <span data-ttu-id="e33f7-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e33f7-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e33f7-149">Properties:</span></span>

|<span data-ttu-id="e33f7-150">Nome</span><span class="sxs-lookup"><span data-stu-id="e33f7-150">Name</span></span>| <span data-ttu-id="e33f7-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="e33f7-151">Type</span></span>| <span data-ttu-id="e33f7-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="e33f7-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e33f7-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-153">String</span></span>|<span data-ttu-id="e33f7-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="e33f7-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e33f7-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-155">String</span></span>|<span data-ttu-id="e33f7-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="e33f7-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e33f7-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e33f7-157">Requirements</span></span>

|<span data-ttu-id="e33f7-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="e33f7-158">Requirement</span></span>| <span data-ttu-id="e33f7-159">Valor</span><span class="sxs-lookup"><span data-stu-id="e33f7-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="e33f7-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e33f7-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e33f7-161">1.0</span><span class="sxs-lookup"><span data-stu-id="e33f7-161">1.0</span></span>|
|[<span data-ttu-id="e33f7-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e33f7-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e33f7-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e33f7-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="e33f7-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="e33f7-164">EventType :String</span></span>

<span data-ttu-id="e33f7-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e33f7-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e33f7-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e33f7-166">Type:</span></span>

*   <span data-ttu-id="e33f7-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e33f7-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e33f7-168">Properties:</span></span>

| <span data-ttu-id="e33f7-169">Nome</span><span class="sxs-lookup"><span data-stu-id="e33f7-169">Name</span></span> | <span data-ttu-id="e33f7-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="e33f7-170">Type</span></span> | <span data-ttu-id="e33f7-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="e33f7-171">Description</span></span> | <span data-ttu-id="e33f7-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="e33f7-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="e33f7-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-173">String</span></span> | <span data-ttu-id="e33f7-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="e33f7-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e33f7-175">1.7</span><span class="sxs-lookup"><span data-stu-id="e33f7-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="e33f7-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-176">String</span></span> | <span data-ttu-id="e33f7-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="e33f7-177">The selected item has changed.</span></span> | <span data-ttu-id="e33f7-178">1,5</span><span class="sxs-lookup"><span data-stu-id="e33f7-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="e33f7-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-179">String</span></span> | <span data-ttu-id="e33f7-180">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="e33f7-180">The selected item has changed.</span></span> | <span data-ttu-id="e33f7-181">Preview</span><span class="sxs-lookup"><span data-stu-id="e33f7-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e33f7-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-182">String</span></span> | <span data-ttu-id="e33f7-183">A lista de destinatários do local de compromisso ou item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="e33f7-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e33f7-184">1.7</span><span class="sxs-lookup"><span data-stu-id="e33f7-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e33f7-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-185">String</span></span> | <span data-ttu-id="e33f7-186">O padrão de recorrência da série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="e33f7-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e33f7-187">1.7</span><span class="sxs-lookup"><span data-stu-id="e33f7-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e33f7-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e33f7-188">Requirements</span></span>

|<span data-ttu-id="e33f7-189">Requisito</span><span class="sxs-lookup"><span data-stu-id="e33f7-189">Requirement</span></span>| <span data-ttu-id="e33f7-190">Valor</span><span class="sxs-lookup"><span data-stu-id="e33f7-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="e33f7-191">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e33f7-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e33f7-192">1,5</span><span class="sxs-lookup"><span data-stu-id="e33f7-192">1.5</span></span> |
|[<span data-ttu-id="e33f7-193">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e33f7-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e33f7-194">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e33f7-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="e33f7-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="e33f7-195">SourceProperty :String</span></span>

<span data-ttu-id="e33f7-196">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="e33f7-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e33f7-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e33f7-197">Type:</span></span>

*   <span data-ttu-id="e33f7-198">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e33f7-199">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e33f7-199">Properties:</span></span>

|<span data-ttu-id="e33f7-200">Nome</span><span class="sxs-lookup"><span data-stu-id="e33f7-200">Name</span></span>| <span data-ttu-id="e33f7-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="e33f7-201">Type</span></span>| <span data-ttu-id="e33f7-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="e33f7-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e33f7-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-203">String</span></span>|<span data-ttu-id="e33f7-204">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e33f7-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e33f7-205">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e33f7-205">String</span></span>|<span data-ttu-id="e33f7-206">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e33f7-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e33f7-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e33f7-207">Requirements</span></span>

|<span data-ttu-id="e33f7-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="e33f7-208">Requirement</span></span>| <span data-ttu-id="e33f7-209">Valor</span><span class="sxs-lookup"><span data-stu-id="e33f7-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="e33f7-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e33f7-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e33f7-211">1.0</span><span class="sxs-lookup"><span data-stu-id="e33f7-211">1.0</span></span>|
|[<span data-ttu-id="e33f7-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e33f7-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e33f7-213">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e33f7-213">Compose or read</span></span>|