# <a name="office"></a><span data-ttu-id="19e8f-101">Office</span><span class="sxs-lookup"><span data-stu-id="19e8f-101">Office</span></span>

<span data-ttu-id="19e8f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="19e8f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="19e8f-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="19e8f-104">Requirements</span></span>

|<span data-ttu-id="19e8f-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="19e8f-105">Requirement</span></span>| <span data-ttu-id="19e8f-106">Valor</span><span class="sxs-lookup"><span data-stu-id="19e8f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="19e8f-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="19e8f-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19e8f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="19e8f-108">1.0</span></span>|
|[<span data-ttu-id="19e8f-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="19e8f-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19e8f-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="19e8f-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="19e8f-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="19e8f-111">Members and methods</span></span>

| <span data-ttu-id="19e8f-112">Membro</span><span class="sxs-lookup"><span data-stu-id="19e8f-112">Member</span></span> | <span data-ttu-id="19e8f-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="19e8f-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="19e8f-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="19e8f-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="19e8f-115">Membro</span><span class="sxs-lookup"><span data-stu-id="19e8f-115">Member</span></span> |
| [<span data-ttu-id="19e8f-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="19e8f-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="19e8f-117">Membro</span><span class="sxs-lookup"><span data-stu-id="19e8f-117">Member</span></span> |
| [<span data-ttu-id="19e8f-118">EventType</span><span class="sxs-lookup"><span data-stu-id="19e8f-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="19e8f-119">Membro</span><span class="sxs-lookup"><span data-stu-id="19e8f-119">Member</span></span> |
| [<span data-ttu-id="19e8f-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="19e8f-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="19e8f-121">Membro</span><span class="sxs-lookup"><span data-stu-id="19e8f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="19e8f-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="19e8f-122">Namespaces</span></span>

<span data-ttu-id="19e8f-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="19e8f-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="19e8f-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="19e8f-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="19e8f-125">Membros</span><span class="sxs-lookup"><span data-stu-id="19e8f-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="19e8f-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="19e8f-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="19e8f-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="19e8f-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="19e8f-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="19e8f-128">Type:</span></span>

*   <span data-ttu-id="19e8f-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19e8f-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="19e8f-130">Properties:</span></span>

|<span data-ttu-id="19e8f-131">Nome</span><span class="sxs-lookup"><span data-stu-id="19e8f-131">Name</span></span>| <span data-ttu-id="19e8f-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="19e8f-132">Type</span></span>| <span data-ttu-id="19e8f-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="19e8f-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="19e8f-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-134">String</span></span>|<span data-ttu-id="19e8f-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="19e8f-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="19e8f-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-136">String</span></span>|<span data-ttu-id="19e8f-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="19e8f-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19e8f-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="19e8f-138">Requirements</span></span>

|<span data-ttu-id="19e8f-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="19e8f-139">Requirement</span></span>| <span data-ttu-id="19e8f-140">Valor</span><span class="sxs-lookup"><span data-stu-id="19e8f-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="19e8f-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="19e8f-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19e8f-142">1.0</span><span class="sxs-lookup"><span data-stu-id="19e8f-142">1.0</span></span>|
|[<span data-ttu-id="19e8f-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="19e8f-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19e8f-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="19e8f-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="19e8f-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="19e8f-145">CoercionType :String</span></span>

<span data-ttu-id="19e8f-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="19e8f-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="19e8f-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="19e8f-147">Type:</span></span>

*   <span data-ttu-id="19e8f-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19e8f-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="19e8f-149">Properties:</span></span>

|<span data-ttu-id="19e8f-150">Nome</span><span class="sxs-lookup"><span data-stu-id="19e8f-150">Name</span></span>| <span data-ttu-id="19e8f-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="19e8f-151">Type</span></span>| <span data-ttu-id="19e8f-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="19e8f-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="19e8f-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-153">String</span></span>|<span data-ttu-id="19e8f-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="19e8f-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="19e8f-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-155">String</span></span>|<span data-ttu-id="19e8f-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="19e8f-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19e8f-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="19e8f-157">Requirements</span></span>

|<span data-ttu-id="19e8f-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="19e8f-158">Requirement</span></span>| <span data-ttu-id="19e8f-159">Valor</span><span class="sxs-lookup"><span data-stu-id="19e8f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="19e8f-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="19e8f-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19e8f-161">1.0</span><span class="sxs-lookup"><span data-stu-id="19e8f-161">1.0</span></span>|
|[<span data-ttu-id="19e8f-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="19e8f-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19e8f-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="19e8f-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="19e8f-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="19e8f-164">EventType :String</span></span>

<span data-ttu-id="19e8f-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="19e8f-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="19e8f-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="19e8f-166">Type:</span></span>

*   <span data-ttu-id="19e8f-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19e8f-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="19e8f-168">Properties:</span></span>

| <span data-ttu-id="19e8f-169">Nome</span><span class="sxs-lookup"><span data-stu-id="19e8f-169">Name</span></span> | <span data-ttu-id="19e8f-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="19e8f-170">Type</span></span> | <span data-ttu-id="19e8f-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="19e8f-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="19e8f-172">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-172">String</span></span> | <span data-ttu-id="19e8f-173">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="19e8f-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="19e8f-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="19e8f-174">Requirements</span></span>

|<span data-ttu-id="19e8f-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="19e8f-175">Requirement</span></span>| <span data-ttu-id="19e8f-176">Valor</span><span class="sxs-lookup"><span data-stu-id="19e8f-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="19e8f-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="19e8f-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19e8f-178">1,5</span><span class="sxs-lookup"><span data-stu-id="19e8f-178">1.5</span></span> |
|[<span data-ttu-id="19e8f-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="19e8f-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19e8f-180">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="19e8f-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="19e8f-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="19e8f-181">SourceProperty :String</span></span>

<span data-ttu-id="19e8f-182">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="19e8f-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="19e8f-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="19e8f-183">Type:</span></span>

*   <span data-ttu-id="19e8f-184">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19e8f-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="19e8f-185">Properties:</span></span>

|<span data-ttu-id="19e8f-186">Nome</span><span class="sxs-lookup"><span data-stu-id="19e8f-186">Name</span></span>| <span data-ttu-id="19e8f-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="19e8f-187">Type</span></span>| <span data-ttu-id="19e8f-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="19e8f-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="19e8f-189">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-189">String</span></span>|<span data-ttu-id="19e8f-190">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="19e8f-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="19e8f-191">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="19e8f-191">String</span></span>|<span data-ttu-id="19e8f-192">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="19e8f-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19e8f-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="19e8f-193">Requirements</span></span>

|<span data-ttu-id="19e8f-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="19e8f-194">Requirement</span></span>| <span data-ttu-id="19e8f-195">Valor</span><span class="sxs-lookup"><span data-stu-id="19e8f-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="19e8f-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="19e8f-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19e8f-197">1.0</span><span class="sxs-lookup"><span data-stu-id="19e8f-197">1.0</span></span>|
|[<span data-ttu-id="19e8f-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="19e8f-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19e8f-199">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="19e8f-199">Compose or read</span></span>|