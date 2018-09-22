 

# <a name="office"></a><span data-ttu-id="b888a-101">Office</span><span class="sxs-lookup"><span data-stu-id="b888a-101">Office</span></span>

<span data-ttu-id="b888a-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b888a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b888a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b888a-104">Requirements</span></span>

|<span data-ttu-id="b888a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b888a-105">Requirement</span></span>| <span data-ttu-id="b888a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b888a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b888a-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b888a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b888a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b888a-108">1.0</span></span>|
|[<span data-ttu-id="b888a-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b888a-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b888a-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b888a-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b888a-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b888a-111">Members and methods</span></span>

| <span data-ttu-id="b888a-112">Membro</span><span class="sxs-lookup"><span data-stu-id="b888a-112">Member</span></span> | <span data-ttu-id="b888a-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="b888a-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b888a-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b888a-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b888a-115">Membro</span><span class="sxs-lookup"><span data-stu-id="b888a-115">Member</span></span> |
| [<span data-ttu-id="b888a-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b888a-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b888a-117">Membro</span><span class="sxs-lookup"><span data-stu-id="b888a-117">Member</span></span> |
| [<span data-ttu-id="b888a-118">EventType</span><span class="sxs-lookup"><span data-stu-id="b888a-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b888a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="b888a-119">Member</span></span> |
| [<span data-ttu-id="b888a-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b888a-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b888a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="b888a-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b888a-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b888a-122">Namespaces</span></span>

<span data-ttu-id="b888a-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="b888a-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b888a-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b888a-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b888a-125">Membros</span><span class="sxs-lookup"><span data-stu-id="b888a-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b888a-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b888a-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="b888a-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="b888a-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b888a-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b888a-128">Type:</span></span>

*   <span data-ttu-id="b888a-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b888a-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b888a-130">Properties:</span></span>

|<span data-ttu-id="b888a-131">Nome</span><span class="sxs-lookup"><span data-stu-id="b888a-131">Name</span></span>| <span data-ttu-id="b888a-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="b888a-132">Type</span></span>| <span data-ttu-id="b888a-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="b888a-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b888a-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-134">String</span></span>|<span data-ttu-id="b888a-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="b888a-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b888a-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-136">String</span></span>|<span data-ttu-id="b888a-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="b888a-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b888a-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b888a-138">Requirements</span></span>

|<span data-ttu-id="b888a-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="b888a-139">Requirement</span></span>| <span data-ttu-id="b888a-140">Valor</span><span class="sxs-lookup"><span data-stu-id="b888a-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="b888a-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b888a-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b888a-142">1.0</span><span class="sxs-lookup"><span data-stu-id="b888a-142">1.0</span></span>|
|[<span data-ttu-id="b888a-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b888a-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b888a-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b888a-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="b888a-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b888a-145">CoercionType :String</span></span>

<span data-ttu-id="b888a-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="b888a-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b888a-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b888a-147">Type:</span></span>

*   <span data-ttu-id="b888a-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b888a-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b888a-149">Properties:</span></span>

|<span data-ttu-id="b888a-150">Nome</span><span class="sxs-lookup"><span data-stu-id="b888a-150">Name</span></span>| <span data-ttu-id="b888a-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="b888a-151">Type</span></span>| <span data-ttu-id="b888a-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="b888a-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b888a-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-153">String</span></span>|<span data-ttu-id="b888a-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="b888a-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b888a-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-155">String</span></span>|<span data-ttu-id="b888a-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="b888a-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b888a-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b888a-157">Requirements</span></span>

|<span data-ttu-id="b888a-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="b888a-158">Requirement</span></span>| <span data-ttu-id="b888a-159">Valor</span><span class="sxs-lookup"><span data-stu-id="b888a-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="b888a-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b888a-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b888a-161">1.0</span><span class="sxs-lookup"><span data-stu-id="b888a-161">1.0</span></span>|
|[<span data-ttu-id="b888a-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b888a-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b888a-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b888a-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="b888a-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b888a-164">EventType :String</span></span>

<span data-ttu-id="b888a-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b888a-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b888a-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b888a-166">Type:</span></span>

*   <span data-ttu-id="b888a-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b888a-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b888a-168">Properties:</span></span>

| <span data-ttu-id="b888a-169">Nome</span><span class="sxs-lookup"><span data-stu-id="b888a-169">Name</span></span> | <span data-ttu-id="b888a-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="b888a-170">Type</span></span> | <span data-ttu-id="b888a-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="b888a-171">Description</span></span> | <span data-ttu-id="b888a-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="b888a-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="b888a-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-173">String</span></span> | <span data-ttu-id="b888a-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b888a-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="b888a-175">1.7</span><span class="sxs-lookup"><span data-stu-id="b888a-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="b888a-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-176">String</span></span> | <span data-ttu-id="b888a-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b888a-177">The selected item has changed.</span></span> | <span data-ttu-id="b888a-178">1,5</span><span class="sxs-lookup"><span data-stu-id="b888a-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="b888a-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-179">String</span></span> | <span data-ttu-id="b888a-180">A lista de destinatários do local de compromisso ou item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b888a-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="b888a-181">1.7</span><span class="sxs-lookup"><span data-stu-id="b888a-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="b888a-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-182">String</span></span> | <span data-ttu-id="b888a-183">O padrão de recorrência da série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b888a-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="b888a-184">1.7</span><span class="sxs-lookup"><span data-stu-id="b888a-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b888a-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b888a-185">Requirements</span></span>

|<span data-ttu-id="b888a-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="b888a-186">Requirement</span></span>| <span data-ttu-id="b888a-187">Valor</span><span class="sxs-lookup"><span data-stu-id="b888a-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="b888a-188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b888a-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b888a-189">1,5</span><span class="sxs-lookup"><span data-stu-id="b888a-189">1.5</span></span> |
|[<span data-ttu-id="b888a-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b888a-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b888a-191">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b888a-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b888a-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b888a-192">SourceProperty :String</span></span>

<span data-ttu-id="b888a-193">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="b888a-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b888a-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b888a-194">Type:</span></span>

*   <span data-ttu-id="b888a-195">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b888a-196">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b888a-196">Properties:</span></span>

|<span data-ttu-id="b888a-197">Nome</span><span class="sxs-lookup"><span data-stu-id="b888a-197">Name</span></span>| <span data-ttu-id="b888a-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="b888a-198">Type</span></span>| <span data-ttu-id="b888a-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="b888a-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b888a-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-200">String</span></span>|<span data-ttu-id="b888a-201">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b888a-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b888a-202">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b888a-202">String</span></span>|<span data-ttu-id="b888a-203">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b888a-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b888a-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b888a-204">Requirements</span></span>

|<span data-ttu-id="b888a-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="b888a-205">Requirement</span></span>| <span data-ttu-id="b888a-206">Valor</span><span class="sxs-lookup"><span data-stu-id="b888a-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="b888a-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b888a-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b888a-208">1.0</span><span class="sxs-lookup"><span data-stu-id="b888a-208">1.0</span></span>|
|[<span data-ttu-id="b888a-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b888a-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b888a-210">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b888a-210">Compose or read</span></span>|