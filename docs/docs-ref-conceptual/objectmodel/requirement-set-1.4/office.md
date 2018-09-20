 

# <a name="office"></a><span data-ttu-id="9a2c1-101">Office</span><span class="sxs-lookup"><span data-stu-id="9a2c1-101">Office</span></span>

<span data-ttu-id="9a2c1-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9a2c1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a2c1-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a2c1-104">Requirements</span></span>

|<span data-ttu-id="9a2c1-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a2c1-105">Requirement</span></span>| <span data-ttu-id="9a2c1-106">Valor</span><span class="sxs-lookup"><span data-stu-id="9a2c1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a2c1-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a2c1-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a2c1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9a2c1-108">1.0</span></span>|
|[<span data-ttu-id="9a2c1-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a2c1-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a2c1-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9a2c1-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="9a2c1-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="9a2c1-111">Namespaces</span></span>

<span data-ttu-id="9a2c1-112">[contexto](Office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9a2c1-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9a2c1-114">Membros</span><span class="sxs-lookup"><span data-stu-id="9a2c1-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9a2c1-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9a2c1-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="9a2c1-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9a2c1-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-117">Type:</span></span>

*   <span data-ttu-id="9a2c1-118">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a2c1-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-119">Properties:</span></span>

|<span data-ttu-id="9a2c1-120">Nome</span><span class="sxs-lookup"><span data-stu-id="9a2c1-120">Name</span></span>| <span data-ttu-id="9a2c1-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a2c1-121">Type</span></span>| <span data-ttu-id="9a2c1-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a2c1-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9a2c1-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-123">String</span></span>|<span data-ttu-id="9a2c1-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9a2c1-125">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-125">String</span></span>|<span data-ttu-id="9a2c1-126">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a2c1-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a2c1-127">Requirements</span></span>

|<span data-ttu-id="9a2c1-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a2c1-128">Requirement</span></span>| <span data-ttu-id="9a2c1-129">Valor</span><span class="sxs-lookup"><span data-stu-id="9a2c1-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a2c1-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a2c1-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a2c1-131">1.0</span><span class="sxs-lookup"><span data-stu-id="9a2c1-131">1.0</span></span>|
|[<span data-ttu-id="9a2c1-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a2c1-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a2c1-133">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9a2c1-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="9a2c1-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9a2c1-134">CoercionType :String</span></span>

<span data-ttu-id="9a2c1-135">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9a2c1-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-136">Type:</span></span>

*   <span data-ttu-id="9a2c1-137">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a2c1-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-138">Properties:</span></span>

|<span data-ttu-id="9a2c1-139">Nome</span><span class="sxs-lookup"><span data-stu-id="9a2c1-139">Name</span></span>| <span data-ttu-id="9a2c1-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a2c1-140">Type</span></span>| <span data-ttu-id="9a2c1-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a2c1-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9a2c1-142">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-142">String</span></span>|<span data-ttu-id="9a2c1-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9a2c1-144">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-144">String</span></span>|<span data-ttu-id="9a2c1-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a2c1-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a2c1-146">Requirements</span></span>

|<span data-ttu-id="9a2c1-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a2c1-147">Requirement</span></span>| <span data-ttu-id="9a2c1-148">Valor</span><span class="sxs-lookup"><span data-stu-id="9a2c1-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a2c1-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a2c1-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a2c1-150">1.0</span><span class="sxs-lookup"><span data-stu-id="9a2c1-150">1.0</span></span>|
|[<span data-ttu-id="9a2c1-151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a2c1-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a2c1-152">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9a2c1-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="9a2c1-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9a2c1-153">SourceProperty :String</span></span>

<span data-ttu-id="9a2c1-154">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9a2c1-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-155">Type:</span></span>

*   <span data-ttu-id="9a2c1-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a2c1-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="9a2c1-157">Properties:</span></span>

|<span data-ttu-id="9a2c1-158">Nome</span><span class="sxs-lookup"><span data-stu-id="9a2c1-158">Name</span></span>| <span data-ttu-id="9a2c1-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a2c1-159">Type</span></span>| <span data-ttu-id="9a2c1-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a2c1-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9a2c1-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-161">String</span></span>|<span data-ttu-id="9a2c1-162">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9a2c1-163">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a2c1-163">String</span></span>|<span data-ttu-id="9a2c1-164">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a2c1-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a2c1-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a2c1-165">Requirements</span></span>

|<span data-ttu-id="9a2c1-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a2c1-166">Requirement</span></span>| <span data-ttu-id="9a2c1-167">Valor</span><span class="sxs-lookup"><span data-stu-id="9a2c1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a2c1-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a2c1-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a2c1-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9a2c1-169">1.0</span></span>|
|[<span data-ttu-id="9a2c1-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a2c1-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a2c1-171">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9a2c1-171">Compose or read</span></span>|