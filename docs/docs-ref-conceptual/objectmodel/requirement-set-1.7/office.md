 

# <a name="office"></a><span data-ttu-id="9305d-101">Office</span><span class="sxs-lookup"><span data-stu-id="9305d-101">Office</span></span>

<span data-ttu-id="9305d-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9305d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9305d-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9305d-104">Requirements</span></span>

|<span data-ttu-id="9305d-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="9305d-105">Requirement</span></span>| <span data-ttu-id="9305d-106">Valor</span><span class="sxs-lookup"><span data-stu-id="9305d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9305d-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9305d-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9305d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9305d-108">1.0</span></span>|
|[<span data-ttu-id="9305d-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9305d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9305d-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9305d-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9305d-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="9305d-111">Members and methods</span></span>

| <span data-ttu-id="9305d-112">Membro</span><span class="sxs-lookup"><span data-stu-id="9305d-112">Member</span></span> | <span data-ttu-id="9305d-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="9305d-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9305d-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9305d-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9305d-115">Membro</span><span class="sxs-lookup"><span data-stu-id="9305d-115">Member</span></span> |
| [<span data-ttu-id="9305d-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9305d-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9305d-117">Membro</span><span class="sxs-lookup"><span data-stu-id="9305d-117">Member</span></span> |
| [<span data-ttu-id="9305d-118">EventType</span><span class="sxs-lookup"><span data-stu-id="9305d-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9305d-119">Membro</span><span class="sxs-lookup"><span data-stu-id="9305d-119">Member</span></span> |
| [<span data-ttu-id="9305d-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9305d-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9305d-121">Membro</span><span class="sxs-lookup"><span data-stu-id="9305d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9305d-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="9305d-122">Namespaces</span></span>

<span data-ttu-id="9305d-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9305d-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9305d-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9305d-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9305d-125">Membros</span><span class="sxs-lookup"><span data-stu-id="9305d-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9305d-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9305d-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="9305d-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="9305d-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9305d-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9305d-128">Type:</span></span>

*   <span data-ttu-id="9305d-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9305d-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9305d-130">Properties:</span></span>

|<span data-ttu-id="9305d-131">Nome</span><span class="sxs-lookup"><span data-stu-id="9305d-131">Name</span></span>| <span data-ttu-id="9305d-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="9305d-132">Type</span></span>| <span data-ttu-id="9305d-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="9305d-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9305d-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-134">String</span></span>|<span data-ttu-id="9305d-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="9305d-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9305d-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-136">String</span></span>|<span data-ttu-id="9305d-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="9305d-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9305d-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9305d-138">Requirements</span></span>

|<span data-ttu-id="9305d-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="9305d-139">Requirement</span></span>| <span data-ttu-id="9305d-140">Valor</span><span class="sxs-lookup"><span data-stu-id="9305d-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="9305d-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9305d-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9305d-142">1.0</span><span class="sxs-lookup"><span data-stu-id="9305d-142">1.0</span></span>|
|[<span data-ttu-id="9305d-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9305d-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9305d-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9305d-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9305d-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9305d-145">CoercionType :String</span></span>

<span data-ttu-id="9305d-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="9305d-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9305d-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9305d-147">Type:</span></span>

*   <span data-ttu-id="9305d-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9305d-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9305d-149">Properties:</span></span>

|<span data-ttu-id="9305d-150">Nome</span><span class="sxs-lookup"><span data-stu-id="9305d-150">Name</span></span>| <span data-ttu-id="9305d-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="9305d-151">Type</span></span>| <span data-ttu-id="9305d-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="9305d-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9305d-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-153">String</span></span>|<span data-ttu-id="9305d-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="9305d-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9305d-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-155">String</span></span>|<span data-ttu-id="9305d-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="9305d-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9305d-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9305d-157">Requirements</span></span>

|<span data-ttu-id="9305d-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="9305d-158">Requirement</span></span>| <span data-ttu-id="9305d-159">Valor</span><span class="sxs-lookup"><span data-stu-id="9305d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9305d-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9305d-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9305d-161">1.0</span><span class="sxs-lookup"><span data-stu-id="9305d-161">1.0</span></span>|
|[<span data-ttu-id="9305d-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9305d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9305d-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9305d-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9305d-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="9305d-164">EventType :String</span></span>

<span data-ttu-id="9305d-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="9305d-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9305d-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9305d-166">Type:</span></span>

*   <span data-ttu-id="9305d-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9305d-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9305d-168">Properties:</span></span>

| <span data-ttu-id="9305d-169">Nome</span><span class="sxs-lookup"><span data-stu-id="9305d-169">Name</span></span> | <span data-ttu-id="9305d-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="9305d-170">Type</span></span> | <span data-ttu-id="9305d-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="9305d-171">Description</span></span> | <span data-ttu-id="9305d-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="9305d-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="9305d-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-173">String</span></span> | <span data-ttu-id="9305d-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="9305d-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9305d-175">1.7</span><span class="sxs-lookup"><span data-stu-id="9305d-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="9305d-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-176">String</span></span> | <span data-ttu-id="9305d-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="9305d-177">The selected item has changed.</span></span> | <span data-ttu-id="9305d-178">1,5</span><span class="sxs-lookup"><span data-stu-id="9305d-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9305d-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-179">String</span></span> | <span data-ttu-id="9305d-180">A lista de destinatários do local de compromisso ou item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="9305d-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9305d-181">1.7</span><span class="sxs-lookup"><span data-stu-id="9305d-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9305d-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-182">String</span></span> | <span data-ttu-id="9305d-183">O padrão de recorrência da série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="9305d-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9305d-184">1.7</span><span class="sxs-lookup"><span data-stu-id="9305d-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9305d-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9305d-185">Requirements</span></span>

|<span data-ttu-id="9305d-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="9305d-186">Requirement</span></span>| <span data-ttu-id="9305d-187">Valor</span><span class="sxs-lookup"><span data-stu-id="9305d-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="9305d-188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9305d-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9305d-189">1,5</span><span class="sxs-lookup"><span data-stu-id="9305d-189">1.5</span></span> |
|[<span data-ttu-id="9305d-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9305d-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9305d-191">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9305d-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9305d-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9305d-192">SourceProperty :String</span></span>

<span data-ttu-id="9305d-193">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="9305d-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9305d-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9305d-194">Type:</span></span>

*   <span data-ttu-id="9305d-195">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9305d-196">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9305d-196">Properties:</span></span>

|<span data-ttu-id="9305d-197">Nome</span><span class="sxs-lookup"><span data-stu-id="9305d-197">Name</span></span>| <span data-ttu-id="9305d-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="9305d-198">Type</span></span>| <span data-ttu-id="9305d-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="9305d-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9305d-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-200">String</span></span>|<span data-ttu-id="9305d-201">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9305d-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9305d-202">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9305d-202">String</span></span>|<span data-ttu-id="9305d-203">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9305d-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9305d-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9305d-204">Requirements</span></span>

|<span data-ttu-id="9305d-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="9305d-205">Requirement</span></span>| <span data-ttu-id="9305d-206">Valor</span><span class="sxs-lookup"><span data-stu-id="9305d-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="9305d-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9305d-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9305d-208">1.0</span><span class="sxs-lookup"><span data-stu-id="9305d-208">1.0</span></span>|
|[<span data-ttu-id="9305d-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9305d-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9305d-210">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9305d-210">Compose or read</span></span>|