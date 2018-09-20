 

# <a name="office"></a><span data-ttu-id="07ffc-101">Office</span><span class="sxs-lookup"><span data-stu-id="07ffc-101">Office</span></span>

<span data-ttu-id="07ffc-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="07ffc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="07ffc-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07ffc-104">Requirements</span></span>

|<span data-ttu-id="07ffc-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="07ffc-105">Requirement</span></span>| <span data-ttu-id="07ffc-106">Valor</span><span class="sxs-lookup"><span data-stu-id="07ffc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="07ffc-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07ffc-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07ffc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="07ffc-108">1.0</span></span>|
|[<span data-ttu-id="07ffc-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07ffc-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07ffc-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07ffc-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="07ffc-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="07ffc-111">Members and methods</span></span>

| <span data-ttu-id="07ffc-112">Membro</span><span class="sxs-lookup"><span data-stu-id="07ffc-112">Member</span></span> | <span data-ttu-id="07ffc-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="07ffc-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="07ffc-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="07ffc-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="07ffc-115">Membro</span><span class="sxs-lookup"><span data-stu-id="07ffc-115">Member</span></span> |
| [<span data-ttu-id="07ffc-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="07ffc-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="07ffc-117">Membro</span><span class="sxs-lookup"><span data-stu-id="07ffc-117">Member</span></span> |
| [<span data-ttu-id="07ffc-118">EventType</span><span class="sxs-lookup"><span data-stu-id="07ffc-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="07ffc-119">Membro</span><span class="sxs-lookup"><span data-stu-id="07ffc-119">Member</span></span> |
| [<span data-ttu-id="07ffc-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="07ffc-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="07ffc-121">Membro</span><span class="sxs-lookup"><span data-stu-id="07ffc-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="07ffc-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="07ffc-122">Namespaces</span></span>

<span data-ttu-id="07ffc-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="07ffc-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="07ffc-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="07ffc-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="07ffc-125">Membros</span><span class="sxs-lookup"><span data-stu-id="07ffc-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="07ffc-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="07ffc-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="07ffc-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="07ffc-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="07ffc-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07ffc-128">Type:</span></span>

*   <span data-ttu-id="07ffc-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="07ffc-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="07ffc-130">Properties:</span></span>

|<span data-ttu-id="07ffc-131">Nome</span><span class="sxs-lookup"><span data-stu-id="07ffc-131">Name</span></span>| <span data-ttu-id="07ffc-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="07ffc-132">Type</span></span>| <span data-ttu-id="07ffc-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="07ffc-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="07ffc-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-134">String</span></span>|<span data-ttu-id="07ffc-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="07ffc-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="07ffc-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-136">String</span></span>|<span data-ttu-id="07ffc-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="07ffc-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07ffc-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07ffc-138">Requirements</span></span>

|<span data-ttu-id="07ffc-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="07ffc-139">Requirement</span></span>| <span data-ttu-id="07ffc-140">Valor</span><span class="sxs-lookup"><span data-stu-id="07ffc-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="07ffc-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07ffc-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07ffc-142">1.0</span><span class="sxs-lookup"><span data-stu-id="07ffc-142">1.0</span></span>|
|[<span data-ttu-id="07ffc-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07ffc-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07ffc-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07ffc-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="07ffc-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="07ffc-145">CoercionType :String</span></span>

<span data-ttu-id="07ffc-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="07ffc-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="07ffc-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07ffc-147">Type:</span></span>

*   <span data-ttu-id="07ffc-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="07ffc-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="07ffc-149">Properties:</span></span>

|<span data-ttu-id="07ffc-150">Nome</span><span class="sxs-lookup"><span data-stu-id="07ffc-150">Name</span></span>| <span data-ttu-id="07ffc-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="07ffc-151">Type</span></span>| <span data-ttu-id="07ffc-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="07ffc-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="07ffc-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-153">String</span></span>|<span data-ttu-id="07ffc-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="07ffc-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="07ffc-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-155">String</span></span>|<span data-ttu-id="07ffc-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="07ffc-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07ffc-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07ffc-157">Requirements</span></span>

|<span data-ttu-id="07ffc-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="07ffc-158">Requirement</span></span>| <span data-ttu-id="07ffc-159">Valor</span><span class="sxs-lookup"><span data-stu-id="07ffc-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="07ffc-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07ffc-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07ffc-161">1.0</span><span class="sxs-lookup"><span data-stu-id="07ffc-161">1.0</span></span>|
|[<span data-ttu-id="07ffc-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07ffc-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07ffc-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07ffc-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="07ffc-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="07ffc-164">EventType :String</span></span>

<span data-ttu-id="07ffc-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="07ffc-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="07ffc-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07ffc-166">Type:</span></span>

*   <span data-ttu-id="07ffc-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="07ffc-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="07ffc-168">Properties:</span></span>

| <span data-ttu-id="07ffc-169">Nome</span><span class="sxs-lookup"><span data-stu-id="07ffc-169">Name</span></span> | <span data-ttu-id="07ffc-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="07ffc-170">Type</span></span> | <span data-ttu-id="07ffc-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="07ffc-171">Description</span></span> | <span data-ttu-id="07ffc-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="07ffc-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="07ffc-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-173">String</span></span> | <span data-ttu-id="07ffc-174">O compromisso de data ou hora da série selecionada foi alterada.</span><span class="sxs-lookup"><span data-stu-id="07ffc-174">The appointment date or time of the selected series has changed.</span></span> | <span data-ttu-id="07ffc-175">Preview</span><span class="sxs-lookup"><span data-stu-id="07ffc-175">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="07ffc-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-176">String</span></span> | <span data-ttu-id="07ffc-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="07ffc-177">The selected item has changed.</span></span> | <span data-ttu-id="07ffc-178">1,5</span><span class="sxs-lookup"><span data-stu-id="07ffc-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="07ffc-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-179">String</span></span> | <span data-ttu-id="07ffc-180">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="07ffc-180">The selected item has changed.</span></span> | <span data-ttu-id="07ffc-181">Preview</span><span class="sxs-lookup"><span data-stu-id="07ffc-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="07ffc-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-182">String</span></span> | <span data-ttu-id="07ffc-183">A lista de destinatários do item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="07ffc-183">The recipient list of the selected item has changed.</span></span> | <span data-ttu-id="07ffc-184">Preview</span><span class="sxs-lookup"><span data-stu-id="07ffc-184">Preview</span></span> |
|`RecurrencePatternChanged`| <span data-ttu-id="07ffc-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-185">String</span></span> | <span data-ttu-id="07ffc-186">O padrão de recorrência da série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="07ffc-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="07ffc-187">Preview</span><span class="sxs-lookup"><span data-stu-id="07ffc-187">Preview</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07ffc-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07ffc-188">Requirements</span></span>

|<span data-ttu-id="07ffc-189">Requisito</span><span class="sxs-lookup"><span data-stu-id="07ffc-189">Requirement</span></span>| <span data-ttu-id="07ffc-190">Valor</span><span class="sxs-lookup"><span data-stu-id="07ffc-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="07ffc-191">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07ffc-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07ffc-192">1,5</span><span class="sxs-lookup"><span data-stu-id="07ffc-192">1.5</span></span> |
|[<span data-ttu-id="07ffc-193">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07ffc-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07ffc-194">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07ffc-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="07ffc-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="07ffc-195">SourceProperty :String</span></span>

<span data-ttu-id="07ffc-196">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="07ffc-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="07ffc-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07ffc-197">Type:</span></span>

*   <span data-ttu-id="07ffc-198">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="07ffc-199">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="07ffc-199">Properties:</span></span>

|<span data-ttu-id="07ffc-200">Nome</span><span class="sxs-lookup"><span data-stu-id="07ffc-200">Name</span></span>| <span data-ttu-id="07ffc-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="07ffc-201">Type</span></span>| <span data-ttu-id="07ffc-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="07ffc-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="07ffc-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-203">String</span></span>|<span data-ttu-id="07ffc-204">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07ffc-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="07ffc-205">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07ffc-205">String</span></span>|<span data-ttu-id="07ffc-206">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07ffc-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07ffc-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07ffc-207">Requirements</span></span>

|<span data-ttu-id="07ffc-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="07ffc-208">Requirement</span></span>| <span data-ttu-id="07ffc-209">Valor</span><span class="sxs-lookup"><span data-stu-id="07ffc-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="07ffc-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07ffc-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07ffc-211">1.0</span><span class="sxs-lookup"><span data-stu-id="07ffc-211">1.0</span></span>|
|[<span data-ttu-id="07ffc-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07ffc-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07ffc-213">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07ffc-213">Compose or read</span></span>|