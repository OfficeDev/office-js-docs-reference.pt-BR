 

# <a name="office"></a><span data-ttu-id="a1116-101">Office</span><span class="sxs-lookup"><span data-stu-id="a1116-101">Office</span></span>

<span data-ttu-id="a1116-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a1116-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1116-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1116-104">Requirements</span></span>

|<span data-ttu-id="a1116-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1116-105">Requirement</span></span>| <span data-ttu-id="a1116-106">Valor</span><span class="sxs-lookup"><span data-stu-id="a1116-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1116-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1116-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1116-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a1116-108">1.0</span></span>|
|[<span data-ttu-id="a1116-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1116-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1116-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a1116-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a1116-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a1116-111">Namespaces</span></span>

<span data-ttu-id="a1116-112">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a1116-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a1116-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a1116-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a1116-114">Membros</span><span class="sxs-lookup"><span data-stu-id="a1116-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a1116-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a1116-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="a1116-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a1116-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a1116-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a1116-117">Type:</span></span>

*   <span data-ttu-id="a1116-118">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1116-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1116-119">Properties:</span></span>

|<span data-ttu-id="a1116-120">Nome</span><span class="sxs-lookup"><span data-stu-id="a1116-120">Name</span></span>| <span data-ttu-id="a1116-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1116-121">Type</span></span>| <span data-ttu-id="a1116-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1116-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a1116-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-123">String</span></span>|<span data-ttu-id="a1116-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a1116-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a1116-125">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-125">String</span></span>|<span data-ttu-id="a1116-126">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a1116-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1116-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1116-127">Requirements</span></span>

|<span data-ttu-id="a1116-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1116-128">Requirement</span></span>| <span data-ttu-id="a1116-129">Valor</span><span class="sxs-lookup"><span data-stu-id="a1116-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1116-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1116-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1116-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a1116-131">1.0</span></span>|
|[<span data-ttu-id="a1116-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1116-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1116-133">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a1116-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="a1116-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a1116-134">CoercionType :String</span></span>

<span data-ttu-id="a1116-135">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="a1116-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a1116-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a1116-136">Type:</span></span>

*   <span data-ttu-id="a1116-137">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1116-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1116-138">Properties:</span></span>

|<span data-ttu-id="a1116-139">Nome</span><span class="sxs-lookup"><span data-stu-id="a1116-139">Name</span></span>| <span data-ttu-id="a1116-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1116-140">Type</span></span>| <span data-ttu-id="a1116-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1116-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a1116-142">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-142">String</span></span>|<span data-ttu-id="a1116-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a1116-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a1116-144">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-144">String</span></span>|<span data-ttu-id="a1116-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a1116-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1116-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1116-146">Requirements</span></span>

|<span data-ttu-id="a1116-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1116-147">Requirement</span></span>| <span data-ttu-id="a1116-148">Valor</span><span class="sxs-lookup"><span data-stu-id="a1116-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1116-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1116-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1116-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a1116-150">1.0</span></span>|
|[<span data-ttu-id="a1116-151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1116-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1116-152">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a1116-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="a1116-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a1116-153">SourceProperty :String</span></span>

<span data-ttu-id="a1116-154">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a1116-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a1116-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a1116-155">Type:</span></span>

*   <span data-ttu-id="a1116-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1116-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1116-157">Properties:</span></span>

|<span data-ttu-id="a1116-158">Nome</span><span class="sxs-lookup"><span data-stu-id="a1116-158">Name</span></span>| <span data-ttu-id="a1116-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1116-159">Type</span></span>| <span data-ttu-id="a1116-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1116-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a1116-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-161">String</span></span>|<span data-ttu-id="a1116-162">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a1116-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a1116-163">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1116-163">String</span></span>|<span data-ttu-id="a1116-164">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a1116-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1116-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1116-165">Requirements</span></span>

|<span data-ttu-id="a1116-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1116-166">Requirement</span></span>| <span data-ttu-id="a1116-167">Valor</span><span class="sxs-lookup"><span data-stu-id="a1116-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1116-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1116-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1116-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a1116-169">1.0</span></span>|
|[<span data-ttu-id="a1116-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1116-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1116-171">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a1116-171">Compose or read</span></span>|