 

# <a name="office"></a><span data-ttu-id="5793a-101">Office</span><span class="sxs-lookup"><span data-stu-id="5793a-101">Office</span></span>

<span data-ttu-id="5793a-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5793a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5793a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5793a-104">Requirements</span></span>

|<span data-ttu-id="5793a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="5793a-105">Requirement</span></span>| <span data-ttu-id="5793a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="5793a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5793a-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5793a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5793a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5793a-108">1.0</span></span>|
|[<span data-ttu-id="5793a-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5793a-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5793a-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5793a-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5793a-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="5793a-111">Members and methods</span></span>

| <span data-ttu-id="5793a-112">Membro</span><span class="sxs-lookup"><span data-stu-id="5793a-112">Member</span></span> | <span data-ttu-id="5793a-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="5793a-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5793a-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5793a-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5793a-115">Membro</span><span class="sxs-lookup"><span data-stu-id="5793a-115">Member</span></span> |
| [<span data-ttu-id="5793a-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5793a-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5793a-117">Membro</span><span class="sxs-lookup"><span data-stu-id="5793a-117">Member</span></span> |
| [<span data-ttu-id="5793a-118">EventType</span><span class="sxs-lookup"><span data-stu-id="5793a-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5793a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="5793a-119">Member</span></span> |
| [<span data-ttu-id="5793a-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5793a-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5793a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="5793a-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5793a-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5793a-122">Namespaces</span></span>

<span data-ttu-id="5793a-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5793a-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5793a-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5793a-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5793a-125">Membros</span><span class="sxs-lookup"><span data-stu-id="5793a-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5793a-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5793a-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="5793a-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5793a-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5793a-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5793a-128">Type:</span></span>

*   <span data-ttu-id="5793a-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5793a-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5793a-130">Properties:</span></span>

|<span data-ttu-id="5793a-131">Nome</span><span class="sxs-lookup"><span data-stu-id="5793a-131">Name</span></span>| <span data-ttu-id="5793a-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="5793a-132">Type</span></span>| <span data-ttu-id="5793a-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="5793a-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5793a-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-134">String</span></span>|<span data-ttu-id="5793a-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5793a-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5793a-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-136">String</span></span>|<span data-ttu-id="5793a-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5793a-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5793a-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5793a-138">Requirements</span></span>

|<span data-ttu-id="5793a-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="5793a-139">Requirement</span></span>| <span data-ttu-id="5793a-140">Valor</span><span class="sxs-lookup"><span data-stu-id="5793a-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="5793a-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5793a-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5793a-142">1.0</span><span class="sxs-lookup"><span data-stu-id="5793a-142">1.0</span></span>|
|[<span data-ttu-id="5793a-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5793a-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5793a-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5793a-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="5793a-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5793a-145">CoercionType :String</span></span>

<span data-ttu-id="5793a-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5793a-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5793a-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5793a-147">Type:</span></span>

*   <span data-ttu-id="5793a-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5793a-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5793a-149">Properties:</span></span>

|<span data-ttu-id="5793a-150">Nome</span><span class="sxs-lookup"><span data-stu-id="5793a-150">Name</span></span>| <span data-ttu-id="5793a-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="5793a-151">Type</span></span>| <span data-ttu-id="5793a-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="5793a-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5793a-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-153">String</span></span>|<span data-ttu-id="5793a-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5793a-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5793a-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-155">String</span></span>|<span data-ttu-id="5793a-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5793a-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5793a-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5793a-157">Requirements</span></span>

|<span data-ttu-id="5793a-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="5793a-158">Requirement</span></span>| <span data-ttu-id="5793a-159">Valor</span><span class="sxs-lookup"><span data-stu-id="5793a-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="5793a-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5793a-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5793a-161">1.0</span><span class="sxs-lookup"><span data-stu-id="5793a-161">1.0</span></span>|
|[<span data-ttu-id="5793a-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5793a-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5793a-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5793a-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="5793a-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="5793a-164">EventType :String</span></span>

<span data-ttu-id="5793a-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="5793a-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5793a-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5793a-166">Type:</span></span>

*   <span data-ttu-id="5793a-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5793a-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5793a-168">Properties:</span></span>

| <span data-ttu-id="5793a-169">Nome</span><span class="sxs-lookup"><span data-stu-id="5793a-169">Name</span></span> | <span data-ttu-id="5793a-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="5793a-170">Type</span></span> | <span data-ttu-id="5793a-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="5793a-171">Description</span></span> | <span data-ttu-id="5793a-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="5793a-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5793a-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-173">String</span></span> | <span data-ttu-id="5793a-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5793a-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5793a-175">1.7</span><span class="sxs-lookup"><span data-stu-id="5793a-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="5793a-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-176">String</span></span> | <span data-ttu-id="5793a-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5793a-177">The selected item has changed.</span></span> | <span data-ttu-id="5793a-178">1,5</span><span class="sxs-lookup"><span data-stu-id="5793a-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="5793a-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-179">String</span></span> | <span data-ttu-id="5793a-180">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5793a-180">The selected item has changed.</span></span> | <span data-ttu-id="5793a-181">Preview</span><span class="sxs-lookup"><span data-stu-id="5793a-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5793a-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-182">String</span></span> | <span data-ttu-id="5793a-183">A lista de destinatários do local de compromisso ou item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5793a-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5793a-184">1.7</span><span class="sxs-lookup"><span data-stu-id="5793a-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5793a-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-185">String</span></span> | <span data-ttu-id="5793a-186">O padrão de recorrência da série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5793a-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5793a-187">1.7</span><span class="sxs-lookup"><span data-stu-id="5793a-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5793a-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5793a-188">Requirements</span></span>

|<span data-ttu-id="5793a-189">Requisito</span><span class="sxs-lookup"><span data-stu-id="5793a-189">Requirement</span></span>| <span data-ttu-id="5793a-190">Valor</span><span class="sxs-lookup"><span data-stu-id="5793a-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="5793a-191">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5793a-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5793a-192">1,5</span><span class="sxs-lookup"><span data-stu-id="5793a-192">1.5</span></span> |
|[<span data-ttu-id="5793a-193">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5793a-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5793a-194">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5793a-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="5793a-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5793a-195">SourceProperty :String</span></span>

<span data-ttu-id="5793a-196">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5793a-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5793a-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5793a-197">Type:</span></span>

*   <span data-ttu-id="5793a-198">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5793a-199">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5793a-199">Properties:</span></span>

|<span data-ttu-id="5793a-200">Nome</span><span class="sxs-lookup"><span data-stu-id="5793a-200">Name</span></span>| <span data-ttu-id="5793a-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="5793a-201">Type</span></span>| <span data-ttu-id="5793a-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="5793a-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5793a-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-203">String</span></span>|<span data-ttu-id="5793a-204">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5793a-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5793a-205">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5793a-205">String</span></span>|<span data-ttu-id="5793a-206">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5793a-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5793a-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5793a-207">Requirements</span></span>

|<span data-ttu-id="5793a-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="5793a-208">Requirement</span></span>| <span data-ttu-id="5793a-209">Valor</span><span class="sxs-lookup"><span data-stu-id="5793a-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="5793a-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5793a-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5793a-211">1.0</span><span class="sxs-lookup"><span data-stu-id="5793a-211">1.0</span></span>|
|[<span data-ttu-id="5793a-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5793a-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5793a-213">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5793a-213">Compose or read</span></span>|