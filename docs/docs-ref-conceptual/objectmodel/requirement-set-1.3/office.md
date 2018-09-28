 

# <a name="office"></a><span data-ttu-id="1737e-101">Office</span><span class="sxs-lookup"><span data-stu-id="1737e-101">Office</span></span>

<span data-ttu-id="1737e-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1737e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1737e-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1737e-104">Requirements</span></span>

|<span data-ttu-id="1737e-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="1737e-105">Requirement</span></span>| <span data-ttu-id="1737e-106">Valor</span><span class="sxs-lookup"><span data-stu-id="1737e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1737e-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1737e-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1737e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1737e-108">1.0</span></span>|
|[<span data-ttu-id="1737e-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1737e-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1737e-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1737e-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="1737e-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="1737e-111">Namespaces</span></span>

<span data-ttu-id="1737e-112">[contexto](office.context.md): fornece interfaces compartilhadas de namespace de contexto a API de suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1737e-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="1737e-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="1737e-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="1737e-114">Membros</span><span class="sxs-lookup"><span data-stu-id="1737e-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="1737e-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="1737e-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="1737e-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1737e-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1737e-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1737e-117">Type:</span></span>

*   <span data-ttu-id="1737e-118">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1737e-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1737e-119">Properties:</span></span>

|<span data-ttu-id="1737e-120">Nome</span><span class="sxs-lookup"><span data-stu-id="1737e-120">Name</span></span>| <span data-ttu-id="1737e-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="1737e-121">Type</span></span>| <span data-ttu-id="1737e-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="1737e-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1737e-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-123">String</span></span>|<span data-ttu-id="1737e-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="1737e-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1737e-125">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-125">String</span></span>|<span data-ttu-id="1737e-126">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="1737e-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1737e-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1737e-127">Requirements</span></span>

|<span data-ttu-id="1737e-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="1737e-128">Requirement</span></span>| <span data-ttu-id="1737e-129">Valor</span><span class="sxs-lookup"><span data-stu-id="1737e-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="1737e-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1737e-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1737e-131">1.0</span><span class="sxs-lookup"><span data-stu-id="1737e-131">1.0</span></span>|
|[<span data-ttu-id="1737e-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1737e-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1737e-133">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1737e-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="1737e-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="1737e-134">CoercionType :String</span></span>

<span data-ttu-id="1737e-135">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="1737e-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1737e-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1737e-136">Type:</span></span>

*   <span data-ttu-id="1737e-137">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1737e-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1737e-138">Properties:</span></span>

|<span data-ttu-id="1737e-139">Nome</span><span class="sxs-lookup"><span data-stu-id="1737e-139">Name</span></span>| <span data-ttu-id="1737e-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="1737e-140">Type</span></span>| <span data-ttu-id="1737e-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="1737e-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1737e-142">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-142">String</span></span>|<span data-ttu-id="1737e-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="1737e-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1737e-144">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-144">String</span></span>|<span data-ttu-id="1737e-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="1737e-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1737e-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1737e-146">Requirements</span></span>

|<span data-ttu-id="1737e-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="1737e-147">Requirement</span></span>| <span data-ttu-id="1737e-148">Valor</span><span class="sxs-lookup"><span data-stu-id="1737e-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="1737e-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1737e-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1737e-150">1.0</span><span class="sxs-lookup"><span data-stu-id="1737e-150">1.0</span></span>|
|[<span data-ttu-id="1737e-151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1737e-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1737e-152">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1737e-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="1737e-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="1737e-153">SourceProperty :String</span></span>

<span data-ttu-id="1737e-154">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="1737e-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1737e-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1737e-155">Type:</span></span>

*   <span data-ttu-id="1737e-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1737e-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1737e-157">Properties:</span></span>

|<span data-ttu-id="1737e-158">Nome</span><span class="sxs-lookup"><span data-stu-id="1737e-158">Name</span></span>| <span data-ttu-id="1737e-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="1737e-159">Type</span></span>| <span data-ttu-id="1737e-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="1737e-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1737e-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-161">String</span></span>|<span data-ttu-id="1737e-162">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1737e-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1737e-163">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1737e-163">String</span></span>|<span data-ttu-id="1737e-164">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1737e-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1737e-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1737e-165">Requirements</span></span>

|<span data-ttu-id="1737e-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="1737e-166">Requirement</span></span>| <span data-ttu-id="1737e-167">Valor</span><span class="sxs-lookup"><span data-stu-id="1737e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1737e-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1737e-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1737e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="1737e-169">1.0</span></span>|
|[<span data-ttu-id="1737e-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1737e-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1737e-171">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1737e-171">Compose or read</span></span>|