 

# <a name="office"></a><span data-ttu-id="40c9e-101">Office</span><span class="sxs-lookup"><span data-stu-id="40c9e-101">Office</span></span>

<span data-ttu-id="40c9e-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="40c9e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="40c9e-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="40c9e-104">Requirements</span></span>

|<span data-ttu-id="40c9e-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="40c9e-105">Requirement</span></span>| <span data-ttu-id="40c9e-106">Valor</span><span class="sxs-lookup"><span data-stu-id="40c9e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="40c9e-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40c9e-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40c9e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="40c9e-108">1.0</span></span>|
|[<span data-ttu-id="40c9e-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="40c9e-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40c9e-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="40c9e-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="40c9e-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="40c9e-111">Members and methods</span></span>

| <span data-ttu-id="40c9e-112">Membro</span><span class="sxs-lookup"><span data-stu-id="40c9e-112">Member</span></span> | <span data-ttu-id="40c9e-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="40c9e-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="40c9e-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="40c9e-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="40c9e-115">Membro</span><span class="sxs-lookup"><span data-stu-id="40c9e-115">Member</span></span> |
| [<span data-ttu-id="40c9e-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="40c9e-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="40c9e-117">Membro</span><span class="sxs-lookup"><span data-stu-id="40c9e-117">Member</span></span> |
| [<span data-ttu-id="40c9e-118">EventType</span><span class="sxs-lookup"><span data-stu-id="40c9e-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="40c9e-119">Membro</span><span class="sxs-lookup"><span data-stu-id="40c9e-119">Member</span></span> |
| [<span data-ttu-id="40c9e-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="40c9e-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="40c9e-121">Membro</span><span class="sxs-lookup"><span data-stu-id="40c9e-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="40c9e-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="40c9e-122">Namespaces</span></span>

<span data-ttu-id="40c9e-123">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="40c9e-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="40c9e-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="40c9e-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="40c9e-125">Membros</span><span class="sxs-lookup"><span data-stu-id="40c9e-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="40c9e-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="40c9e-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="40c9e-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="40c9e-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="40c9e-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="40c9e-128">Type:</span></span>

*   <span data-ttu-id="40c9e-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="40c9e-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="40c9e-130">Properties:</span></span>

|<span data-ttu-id="40c9e-131">Nome</span><span class="sxs-lookup"><span data-stu-id="40c9e-131">Name</span></span>| <span data-ttu-id="40c9e-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="40c9e-132">Type</span></span>| <span data-ttu-id="40c9e-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="40c9e-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="40c9e-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-134">String</span></span>|<span data-ttu-id="40c9e-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="40c9e-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="40c9e-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-136">String</span></span>|<span data-ttu-id="40c9e-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="40c9e-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40c9e-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="40c9e-138">Requirements</span></span>

|<span data-ttu-id="40c9e-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="40c9e-139">Requirement</span></span>| <span data-ttu-id="40c9e-140">Valor</span><span class="sxs-lookup"><span data-stu-id="40c9e-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="40c9e-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40c9e-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40c9e-142">1.0</span><span class="sxs-lookup"><span data-stu-id="40c9e-142">1.0</span></span>|
|[<span data-ttu-id="40c9e-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="40c9e-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40c9e-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="40c9e-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="40c9e-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="40c9e-145">CoercionType :String</span></span>

<span data-ttu-id="40c9e-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="40c9e-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="40c9e-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="40c9e-147">Type:</span></span>

*   <span data-ttu-id="40c9e-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="40c9e-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="40c9e-149">Properties:</span></span>

|<span data-ttu-id="40c9e-150">Nome</span><span class="sxs-lookup"><span data-stu-id="40c9e-150">Name</span></span>| <span data-ttu-id="40c9e-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="40c9e-151">Type</span></span>| <span data-ttu-id="40c9e-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="40c9e-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="40c9e-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-153">String</span></span>|<span data-ttu-id="40c9e-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="40c9e-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="40c9e-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-155">String</span></span>|<span data-ttu-id="40c9e-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="40c9e-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40c9e-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="40c9e-157">Requirements</span></span>

|<span data-ttu-id="40c9e-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="40c9e-158">Requirement</span></span>| <span data-ttu-id="40c9e-159">Valor</span><span class="sxs-lookup"><span data-stu-id="40c9e-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="40c9e-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40c9e-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40c9e-161">1.0</span><span class="sxs-lookup"><span data-stu-id="40c9e-161">1.0</span></span>|
|[<span data-ttu-id="40c9e-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="40c9e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40c9e-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="40c9e-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="40c9e-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="40c9e-164">EventType :String</span></span>

<span data-ttu-id="40c9e-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="40c9e-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="40c9e-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="40c9e-166">Type:</span></span>

*   <span data-ttu-id="40c9e-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="40c9e-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="40c9e-168">Properties:</span></span>

| <span data-ttu-id="40c9e-169">Nome</span><span class="sxs-lookup"><span data-stu-id="40c9e-169">Name</span></span> | <span data-ttu-id="40c9e-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="40c9e-170">Type</span></span> | <span data-ttu-id="40c9e-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="40c9e-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="40c9e-172">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-172">String</span></span> | <span data-ttu-id="40c9e-173">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="40c9e-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="40c9e-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="40c9e-174">Requirements</span></span>

|<span data-ttu-id="40c9e-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="40c9e-175">Requirement</span></span>| <span data-ttu-id="40c9e-176">Valor</span><span class="sxs-lookup"><span data-stu-id="40c9e-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="40c9e-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40c9e-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40c9e-178">1,5</span><span class="sxs-lookup"><span data-stu-id="40c9e-178">1.5</span></span> |
|[<span data-ttu-id="40c9e-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="40c9e-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40c9e-180">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="40c9e-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="40c9e-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="40c9e-181">SourceProperty :String</span></span>

<span data-ttu-id="40c9e-182">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="40c9e-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="40c9e-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="40c9e-183">Type:</span></span>

*   <span data-ttu-id="40c9e-184">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="40c9e-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="40c9e-185">Properties:</span></span>

|<span data-ttu-id="40c9e-186">Nome</span><span class="sxs-lookup"><span data-stu-id="40c9e-186">Name</span></span>| <span data-ttu-id="40c9e-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="40c9e-187">Type</span></span>| <span data-ttu-id="40c9e-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="40c9e-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="40c9e-189">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-189">String</span></span>|<span data-ttu-id="40c9e-190">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="40c9e-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="40c9e-191">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="40c9e-191">String</span></span>|<span data-ttu-id="40c9e-192">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="40c9e-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="40c9e-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="40c9e-193">Requirements</span></span>

|<span data-ttu-id="40c9e-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="40c9e-194">Requirement</span></span>| <span data-ttu-id="40c9e-195">Valor</span><span class="sxs-lookup"><span data-stu-id="40c9e-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="40c9e-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40c9e-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40c9e-197">1.0</span><span class="sxs-lookup"><span data-stu-id="40c9e-197">1.0</span></span>|
|[<span data-ttu-id="40c9e-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="40c9e-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40c9e-199">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="40c9e-199">Compose or read</span></span>|