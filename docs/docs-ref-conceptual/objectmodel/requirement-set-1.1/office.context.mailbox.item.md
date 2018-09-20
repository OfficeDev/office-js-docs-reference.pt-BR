
# <a name="item"></a><span data-ttu-id="9c730-101">item</span><span class="sxs-lookup"><span data-stu-id="9c730-101">item</span></span>

### <span data-ttu-id="9c730-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="9c730-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9c730-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9c730-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-106">Requirements</span></span>

|<span data-ttu-id="9c730-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-107">Requirement</span></span>| <span data-ttu-id="9c730-108">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-110">1.0</span></span>|
|[<span data-ttu-id="9c730-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="9c730-112">Restricted</span></span>|
|[<span data-ttu-id="9c730-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="9c730-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-115">Example</span></span>

<span data-ttu-id="9c730-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c730-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="9c730-117">Membros</span><span class="sxs-lookup"><span data-stu-id="9c730-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="9c730-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9c730-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="9c730-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-121">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="9c730-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9c730-122">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9c730-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-123">Type:</span></span>

*   <span data-ttu-id="9c730-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9c730-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-125">Requirements</span></span>

|<span data-ttu-id="9c730-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-126">Requirement</span></span>| <span data-ttu-id="9c730-127">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-128">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-129">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-129">1.0</span></span>|
|[<span data-ttu-id="9c730-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-131">ReadItem</span></span>|
|[<span data-ttu-id="9c730-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-133">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-134">Example</span></span>

<span data-ttu-id="9c730-135">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9c730-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9c730-137">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9c730-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9c730-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9c730-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-139">Type:</span></span>

*   [<span data-ttu-id="9c730-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="9c730-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9c730-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-141">Requirements</span></span>

|<span data-ttu-id="9c730-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-142">Requirement</span></span>| <span data-ttu-id="9c730-143">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-145">1.1</span><span class="sxs-lookup"><span data-stu-id="9c730-145">1.1</span></span>|
|[<span data-ttu-id="9c730-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-147">ReadItem</span></span>|
|[<span data-ttu-id="9c730-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-149">Composição</span><span class="sxs-lookup"><span data-stu-id="9c730-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="9c730-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="9c730-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="9c730-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="9c730-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-153">Type:</span></span>

*   [<span data-ttu-id="9c730-154">Corpo</span><span class="sxs-lookup"><span data-stu-id="9c730-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="9c730-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-155">Requirements</span></span>

|<span data-ttu-id="9c730-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-156">Requirement</span></span>| <span data-ttu-id="9c730-157">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9c730-159">1.1</span></span>|
|[<span data-ttu-id="9c730-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-161">ReadItem</span></span>|
|[<span data-ttu-id="9c730-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9c730-164">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9c730-165">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9c730-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9c730-166">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-167">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-167">Read mode</span></span>

<span data-ttu-id="9c730-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9c730-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-170">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-170">Compose mode</span></span>

<span data-ttu-id="9c730-171">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9c730-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-172">Type:</span></span>

*   <span data-ttu-id="9c730-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-174">Requirements</span></span>

|<span data-ttu-id="9c730-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-175">Requirement</span></span>| <span data-ttu-id="9c730-176">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-178">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-178">1.0</span></span>|
|[<span data-ttu-id="9c730-179">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-180">ReadItem</span></span>|
|[<span data-ttu-id="9c730-181">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-182">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9c730-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9c730-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="9c730-185">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="9c730-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9c730-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="9c730-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9c730-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="9c730-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-190">Type:</span></span>

*   <span data-ttu-id="9c730-191">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-192">Requirements</span></span>

|<span data-ttu-id="9c730-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-193">Requirement</span></span>| <span data-ttu-id="9c730-194">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-196">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-196">1.0</span></span>|
|[<span data-ttu-id="9c730-197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-198">ReadItem</span></span>|
|[<span data-ttu-id="9c730-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-200">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="9c730-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9c730-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="9c730-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-204">Type:</span></span>

*   <span data-ttu-id="9c730-205">Data</span><span class="sxs-lookup"><span data-stu-id="9c730-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-206">Requirements</span></span>

|<span data-ttu-id="9c730-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-207">Requirement</span></span>| <span data-ttu-id="9c730-208">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-210">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-210">1.0</span></span>|
|[<span data-ttu-id="9c730-211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-212">ReadItem</span></span>|
|[<span data-ttu-id="9c730-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-214">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9c730-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9c730-216">dateTimeModified :Date</span></span>

<span data-ttu-id="9c730-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-219">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-220">Type:</span></span>

*   <span data-ttu-id="9c730-221">Data</span><span class="sxs-lookup"><span data-stu-id="9c730-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-222">Requirements</span></span>

|<span data-ttu-id="9c730-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-223">Requirement</span></span>| <span data-ttu-id="9c730-224">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-226">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-226">1.0</span></span>|
|[<span data-ttu-id="9c730-227">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-228">ReadItem</span></span>|
|[<span data-ttu-id="9c730-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-230">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-231">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9c730-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9c730-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9c730-233">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="9c730-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9c730-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="9c730-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-236">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-236">Read mode</span></span>

<span data-ttu-id="9c730-237">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="9c730-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-238">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-238">Compose mode</span></span>

<span data-ttu-id="9c730-239">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9c730-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9c730-240">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="9c730-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-241">Type:</span></span>

*   <span data-ttu-id="9c730-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9c730-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-243">Requirements</span></span>

|<span data-ttu-id="9c730-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-244">Requirement</span></span>| <span data-ttu-id="9c730-245">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-246">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-247">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-247">1.0</span></span>|
|[<span data-ttu-id="9c730-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-249">ReadItem</span></span>|
|[<span data-ttu-id="9c730-250">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-251">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-252">Example</span></span>

<span data-ttu-id="9c730-253">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9c730-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9c730-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9c730-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9c730-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9c730-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="9c730-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-259">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9c730-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-260">Type:</span></span>

*   [<span data-ttu-id="9c730-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9c730-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9c730-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-262">Requirements</span></span>

|<span data-ttu-id="9c730-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-263">Requirement</span></span>| <span data-ttu-id="9c730-264">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-265">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-266">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-266">1.0</span></span>|
|[<span data-ttu-id="9c730-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-268">ReadItem</span></span>|
|[<span data-ttu-id="9c730-269">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-270">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="9c730-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9c730-271">internetMessageId :String</span></span>

<span data-ttu-id="9c730-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-274">Type:</span></span>

*   <span data-ttu-id="9c730-275">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-276">Requirements</span></span>

|<span data-ttu-id="9c730-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-277">Requirement</span></span>| <span data-ttu-id="9c730-278">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-280">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-280">1.0</span></span>|
|[<span data-ttu-id="9c730-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-282">ReadItem</span></span>|
|[<span data-ttu-id="9c730-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-284">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9c730-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9c730-286">itemClass :String</span></span>

<span data-ttu-id="9c730-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9c730-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9c730-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-291">Type</span></span> | <span data-ttu-id="9c730-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-292">Description</span></span> | <span data-ttu-id="9c730-293">classe de item</span><span class="sxs-lookup"><span data-stu-id="9c730-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9c730-294">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="9c730-294">Appointment items</span></span> | <span data-ttu-id="9c730-295">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="9c730-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="9c730-296">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="9c730-296">Message items</span></span> | <span data-ttu-id="9c730-297">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="9c730-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9c730-298">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="9c730-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-299">Type:</span></span>

*   <span data-ttu-id="9c730-300">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-301">Requirements</span></span>

|<span data-ttu-id="9c730-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-302">Requirement</span></span>| <span data-ttu-id="9c730-303">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-305">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-305">1.0</span></span>|
|[<span data-ttu-id="9c730-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-307">ReadItem</span></span>|
|[<span data-ttu-id="9c730-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-309">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9c730-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9c730-311">(nullable) itemId :String</span></span>

<span data-ttu-id="9c730-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-314">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c730-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9c730-315">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c730-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9c730-316">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do requisito definido 1.3.</span><span class="sxs-lookup"><span data-stu-id="9c730-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9c730-317">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9c730-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-318">Type:</span></span>

*   <span data-ttu-id="9c730-319">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-320">Requirements</span></span>

|<span data-ttu-id="9c730-321">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-321">Requirement</span></span>| <span data-ttu-id="9c730-322">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-323">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-324">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-324">1.0</span></span>|
|[<span data-ttu-id="9c730-325">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-326">ReadItem</span></span>|
|[<span data-ttu-id="9c730-327">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-328">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-329">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-329">Example</span></span>

<span data-ttu-id="9c730-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="9c730-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="9c730-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9c730-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9c730-333">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="9c730-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9c730-334">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-335">Type:</span></span>

*   [<span data-ttu-id="9c730-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9c730-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9c730-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-337">Requirements</span></span>

|<span data-ttu-id="9c730-338">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-338">Requirement</span></span>| <span data-ttu-id="9c730-339">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-340">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-341">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-341">1.0</span></span>|
|[<span data-ttu-id="9c730-342">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-343">ReadItem</span></span>|
|[<span data-ttu-id="9c730-344">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-345">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="9c730-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9c730-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="9c730-348">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-349">Read mode</span></span>

<span data-ttu-id="9c730-350">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-351">Compose mode</span></span>

<span data-ttu-id="9c730-352">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-353">Type:</span></span>

*   <span data-ttu-id="9c730-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9c730-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-355">Requirements</span></span>

|<span data-ttu-id="9c730-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-356">Requirement</span></span>| <span data-ttu-id="9c730-357">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-358">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-359">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-359">1.0</span></span>|
|[<span data-ttu-id="9c730-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-361">ReadItem</span></span>|
|[<span data-ttu-id="9c730-362">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-363">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-364">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9c730-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9c730-365">normalizedSubject :String</span></span>

<span data-ttu-id="9c730-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9c730-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="9c730-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-370">Type:</span></span>

*   <span data-ttu-id="9c730-371">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-372">Requirements</span></span>

|<span data-ttu-id="9c730-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-373">Requirement</span></span>| <span data-ttu-id="9c730-374">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-376">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-376">1.0</span></span>|
|[<span data-ttu-id="9c730-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-378">ReadItem</span></span>|
|[<span data-ttu-id="9c730-379">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-380">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-381">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9c730-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9c730-383">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="9c730-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9c730-384">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-385">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-385">Read mode</span></span>

<span data-ttu-id="9c730-386">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="9c730-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-387">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-387">Compose mode</span></span>

<span data-ttu-id="9c730-388">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="9c730-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-389">Type:</span></span>

*   <span data-ttu-id="9c730-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-391">Requirements</span></span>

|<span data-ttu-id="9c730-392">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-392">Requirement</span></span>| <span data-ttu-id="9c730-393">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-394">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-395">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-395">1.0</span></span>|
|[<span data-ttu-id="9c730-396">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-397">ReadItem</span></span>|
|[<span data-ttu-id="9c730-398">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-399">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-400">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9c730-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9c730-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9c730-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-404">Type:</span></span>

*   [<span data-ttu-id="9c730-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9c730-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9c730-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-406">Requirements</span></span>

|<span data-ttu-id="9c730-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-407">Requirement</span></span>| <span data-ttu-id="9c730-408">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-410">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-410">1.0</span></span>|
|[<span data-ttu-id="9c730-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-412">ReadItem</span></span>|
|[<span data-ttu-id="9c730-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-414">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9c730-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9c730-417">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="9c730-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9c730-418">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-419">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-419">Read mode</span></span>

<span data-ttu-id="9c730-420">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="9c730-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-421">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-421">Compose mode</span></span>

<span data-ttu-id="9c730-422">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="9c730-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-423">Type:</span></span>

*   <span data-ttu-id="9c730-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-425">Requirements</span></span>

|<span data-ttu-id="9c730-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-426">Requirement</span></span>| <span data-ttu-id="9c730-427">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-429">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-429">1.0</span></span>|
|[<span data-ttu-id="9c730-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-431">ReadItem</span></span>|
|[<span data-ttu-id="9c730-432">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-433">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9c730-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9c730-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9c730-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9c730-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9c730-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="9c730-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-440">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9c730-440">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-441">Type:</span></span>

*   [<span data-ttu-id="9c730-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9c730-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9c730-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-443">Requirements</span></span>

|<span data-ttu-id="9c730-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-444">Requirement</span></span>| <span data-ttu-id="9c730-445">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-447">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-447">1.0</span></span>|
|[<span data-ttu-id="9c730-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-449">ReadItem</span></span>|
|[<span data-ttu-id="9c730-450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-451">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9c730-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9c730-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9c730-454">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="9c730-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9c730-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="9c730-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-457">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-457">Read mode</span></span>

<span data-ttu-id="9c730-458">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="9c730-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-459">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-459">Compose mode</span></span>

<span data-ttu-id="9c730-460">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9c730-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9c730-461">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="9c730-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-462">Type:</span></span>

*   <span data-ttu-id="9c730-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9c730-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-464">Requirements</span></span>

|<span data-ttu-id="9c730-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-465">Requirement</span></span>| <span data-ttu-id="9c730-466">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-468">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-468">1.0</span></span>|
|[<span data-ttu-id="9c730-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-470">ReadItem</span></span>|
|[<span data-ttu-id="9c730-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-472">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-473">Example</span></span>

<span data-ttu-id="9c730-474">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9c730-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="9c730-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9c730-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="9c730-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="9c730-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9c730-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="9c730-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-478">Read mode</span></span>

<span data-ttu-id="9c730-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9c730-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="9c730-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-481">Compose mode</span></span>

<span data-ttu-id="9c730-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="9c730-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9c730-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-483">Type:</span></span>

*   <span data-ttu-id="9c730-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9c730-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-485">Requirements</span></span>

|<span data-ttu-id="9c730-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-486">Requirement</span></span>| <span data-ttu-id="9c730-487">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-489">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-489">1.0</span></span>|
|[<span data-ttu-id="9c730-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-491">ReadItem</span></span>|
|[<span data-ttu-id="9c730-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-493">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9c730-494">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9c730-495">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9c730-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9c730-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9c730-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-497">Read mode</span></span>

<span data-ttu-id="9c730-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9c730-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9c730-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9c730-500">Compose mode</span></span>

<span data-ttu-id="9c730-501">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9c730-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9c730-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9c730-502">Type:</span></span>

*   <span data-ttu-id="9c730-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9c730-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-504">Requirements</span></span>

|<span data-ttu-id="9c730-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-505">Requirement</span></span>| <span data-ttu-id="9c730-506">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-508">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-508">1.0</span></span>|
|[<span data-ttu-id="9c730-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-510">ReadItem</span></span>|
|[<span data-ttu-id="9c730-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-512">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="9c730-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="9c730-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9c730-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9c730-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9c730-516">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="9c730-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9c730-517">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="9c730-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9c730-518">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9c730-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-519">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-519">Parameters:</span></span>

|<span data-ttu-id="9c730-520">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-520">Name</span></span>| <span data-ttu-id="9c730-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-521">Type</span></span>| <span data-ttu-id="9c730-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="9c730-522">Attributes</span></span>| <span data-ttu-id="9c730-523">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9c730-524">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-524">String</span></span>||<span data-ttu-id="9c730-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9c730-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9c730-527">String</span><span class="sxs-lookup"><span data-stu-id="9c730-527">String</span></span>||<span data-ttu-id="9c730-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9c730-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9c730-530">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-530">Object</span></span>| <span data-ttu-id="9c730-531">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-531">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-532">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9c730-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-533">Object</span></span>| <span data-ttu-id="9c730-534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-534">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-535">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9c730-536">function</span><span class="sxs-lookup"><span data-stu-id="9c730-536">function</span></span>| <span data-ttu-id="9c730-537">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-537">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-538">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9c730-539">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c730-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9c730-540">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="9c730-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c730-541">Erros</span><span class="sxs-lookup"><span data-stu-id="9c730-541">Errors</span></span>

| <span data-ttu-id="9c730-542">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9c730-542">Error code</span></span> | <span data-ttu-id="9c730-543">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9c730-544">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="9c730-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9c730-545">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="9c730-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9c730-546">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="9c730-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c730-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-547">Requirements</span></span>

|<span data-ttu-id="9c730-548">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-548">Requirement</span></span>| <span data-ttu-id="9c730-549">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-550">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-551">1.1</span><span class="sxs-lookup"><span data-stu-id="9c730-551">1.1</span></span>|
|[<span data-ttu-id="9c730-552">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9c730-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="9c730-554">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-555">Composição</span><span class="sxs-lookup"><span data-stu-id="9c730-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-556">Example</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9c730-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9c730-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9c730-558">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9c730-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="9c730-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9c730-562">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9c730-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9c730-563">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="9c730-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-564">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-564">Parameters:</span></span>

|<span data-ttu-id="9c730-565">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-565">Name</span></span>| <span data-ttu-id="9c730-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-566">Type</span></span>| <span data-ttu-id="9c730-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="9c730-567">Attributes</span></span>| <span data-ttu-id="9c730-568">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9c730-569">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-569">String</span></span>||<span data-ttu-id="9c730-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9c730-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9c730-572">String</span><span class="sxs-lookup"><span data-stu-id="9c730-572">String</span></span>||<span data-ttu-id="9c730-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9c730-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9c730-575">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-575">Object</span></span>| <span data-ttu-id="9c730-576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-576">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-577">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9c730-578">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-578">Object</span></span>| <span data-ttu-id="9c730-579">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-579">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-580">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9c730-581">function</span><span class="sxs-lookup"><span data-stu-id="9c730-581">function</span></span>| <span data-ttu-id="9c730-582">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-582">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-583">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9c730-584">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c730-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9c730-585">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="9c730-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c730-586">Erros</span><span class="sxs-lookup"><span data-stu-id="9c730-586">Errors</span></span>

| <span data-ttu-id="9c730-587">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9c730-587">Error code</span></span> | <span data-ttu-id="9c730-588">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9c730-589">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="9c730-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c730-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-590">Requirements</span></span>

|<span data-ttu-id="9c730-591">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-591">Requirement</span></span>| <span data-ttu-id="9c730-592">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-593">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-594">1.1</span><span class="sxs-lookup"><span data-stu-id="9c730-594">1.1</span></span>|
|[<span data-ttu-id="9c730-595">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9c730-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="9c730-597">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-598">Composição</span><span class="sxs-lookup"><span data-stu-id="9c730-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-599">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-599">Example</span></span>

<span data-ttu-id="9c730-600">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9c730-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="9c730-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9c730-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="9c730-602">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="9c730-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-603">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c730-604">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="9c730-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9c730-605">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="9c730-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-606">A capacidade de incluir anexos na chamada a `displayReplyAllForm` não é suportado no conjunto de requisito 1.1.</span><span class="sxs-lookup"><span data-stu-id="9c730-606">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9c730-607">O suporte a anexos foi adicionado a `displayReplyAllForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="9c730-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-608">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-608">Parameters:</span></span>

|<span data-ttu-id="9c730-609">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-609">Name</span></span>| <span data-ttu-id="9c730-610">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-610">Type</span></span>| <span data-ttu-id="9c730-611">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9c730-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9c730-612">String &#124; Object</span></span>| |<span data-ttu-id="9c730-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9c730-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9c730-615">**OU**</span><span class="sxs-lookup"><span data-stu-id="9c730-615">**OR**</span></span><br/><span data-ttu-id="9c730-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9c730-618">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-618">String</span></span> | <span data-ttu-id="9c730-619">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-619">&lt;optional&gt;</span></span> | <span data-ttu-id="9c730-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9c730-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9c730-622">function</span><span class="sxs-lookup"><span data-stu-id="9c730-622">function</span></span> | <span data-ttu-id="9c730-623">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-623">&lt;optional&gt;</span></span> | <span data-ttu-id="9c730-624">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c730-625">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-625">Requirements</span></span>

|<span data-ttu-id="9c730-626">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-626">Requirement</span></span>| <span data-ttu-id="9c730-627">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-628">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-628">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-629">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-629">1.0</span></span>|
|[<span data-ttu-id="9c730-630">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-631">ReadItem</span></span>|
|[<span data-ttu-id="9c730-632">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-633">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9c730-634">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9c730-634">Examples</span></span>

<span data-ttu-id="9c730-635">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9c730-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9c730-636">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="9c730-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9c730-637">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="9c730-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9c730-638">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="9c730-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9c730-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="9c730-640">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="9c730-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-641">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-641">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c730-642">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="9c730-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9c730-643">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="9c730-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-644">A capacidade de incluir anexos na chamada a `displayReplyForm` não é suportado no conjunto de requisito 1.1.</span><span class="sxs-lookup"><span data-stu-id="9c730-644">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9c730-645">O suporte a anexos foi adicionado a `displayReplyForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="9c730-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-646">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-646">Parameters:</span></span>

|<span data-ttu-id="9c730-647">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-647">Name</span></span>| <span data-ttu-id="9c730-648">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-648">Type</span></span>| <span data-ttu-id="9c730-649">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9c730-650">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9c730-650">String &#124; Object</span></span>| | <span data-ttu-id="9c730-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9c730-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9c730-653">**OU**</span><span class="sxs-lookup"><span data-stu-id="9c730-653">**OR**</span></span><br/><span data-ttu-id="9c730-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9c730-656">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-656">String</span></span> | <span data-ttu-id="9c730-657">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-657">&lt;optional&gt;</span></span> | <span data-ttu-id="9c730-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9c730-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9c730-660">function</span><span class="sxs-lookup"><span data-stu-id="9c730-660">function</span></span> | <span data-ttu-id="9c730-661">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-661">&lt;optional&gt;</span></span> | <span data-ttu-id="9c730-662">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c730-663">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-663">Requirements</span></span>

|<span data-ttu-id="9c730-664">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-664">Requirement</span></span>| <span data-ttu-id="9c730-665">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-666">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-666">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-667">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-667">1.0</span></span>|
|[<span data-ttu-id="9c730-668">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-669">ReadItem</span></span>|
|[<span data-ttu-id="9c730-670">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-671">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9c730-672">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9c730-672">Examples</span></span>

<span data-ttu-id="9c730-673">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9c730-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9c730-674">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="9c730-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9c730-675">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="9c730-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9c730-676">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="9c730-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9c730-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="9c730-678">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9c730-678">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-679">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-679">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-680">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-680">Requirements</span></span>

|<span data-ttu-id="9c730-681">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-681">Requirement</span></span>| <span data-ttu-id="9c730-682">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-683">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-683">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-684">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-684">1.0</span></span>|
|[<span data-ttu-id="9c730-685">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-686">ReadItem</span></span>|
|[<span data-ttu-id="9c730-687">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-688">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c730-689">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9c730-689">Returns:</span></span>

<span data-ttu-id="9c730-690">Tipo: [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9c730-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9c730-691">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-691">Example</span></span>

<span data-ttu-id="9c730-692">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-692">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9c730-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9c730-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9c730-694">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9c730-694">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-695">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-695">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-696">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-696">Parameters:</span></span>

|<span data-ttu-id="9c730-697">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-697">Name</span></span>| <span data-ttu-id="9c730-698">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-698">Type</span></span>| <span data-ttu-id="9c730-699">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9c730-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9c730-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="9c730-701">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="9c730-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c730-702">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-702">Requirements</span></span>

|<span data-ttu-id="9c730-703">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-703">Requirement</span></span>| <span data-ttu-id="9c730-704">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-705">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-705">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-706">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-706">1.0</span></span>|
|[<span data-ttu-id="9c730-707">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-708">Restrito</span><span class="sxs-lookup"><span data-stu-id="9c730-708">Restricted</span></span>|
|[<span data-ttu-id="9c730-709">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-710">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c730-711">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9c730-711">Returns:</span></span>

<span data-ttu-id="9c730-712">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="9c730-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9c730-713">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="9c730-713">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9c730-714">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9c730-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9c730-715">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9c730-716">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="9c730-716">Value of `entityType`</span></span> | <span data-ttu-id="9c730-717">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="9c730-717">Type of objects in returned array</span></span> | <span data-ttu-id="9c730-718">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="9c730-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9c730-719">String</span><span class="sxs-lookup"><span data-stu-id="9c730-719">String</span></span> | <span data-ttu-id="9c730-720">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9c730-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9c730-721">Contato</span><span class="sxs-lookup"><span data-stu-id="9c730-721">Contact</span></span> | <span data-ttu-id="9c730-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9c730-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9c730-723">String</span><span class="sxs-lookup"><span data-stu-id="9c730-723">String</span></span> | <span data-ttu-id="9c730-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9c730-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9c730-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9c730-725">MeetingSuggestion</span></span> | <span data-ttu-id="9c730-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9c730-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9c730-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9c730-727">PhoneNumber</span></span> | <span data-ttu-id="9c730-728">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9c730-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9c730-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9c730-729">TaskSuggestion</span></span> | <span data-ttu-id="9c730-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9c730-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9c730-731">String</span><span class="sxs-lookup"><span data-stu-id="9c730-731">String</span></span> | <span data-ttu-id="9c730-732">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9c730-732">**Restricted**</span></span> |

<span data-ttu-id="9c730-733">Tipo:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9c730-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="9c730-734">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-734">Example</span></span>

<span data-ttu-id="9c730-735">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9c730-735">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9c730-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9c730-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9c730-737">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9c730-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-738">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-738">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c730-739">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="9c730-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-740">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-740">Parameters:</span></span>

|<span data-ttu-id="9c730-741">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-741">Name</span></span>| <span data-ttu-id="9c730-742">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-742">Type</span></span>| <span data-ttu-id="9c730-743">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9c730-744">String</span><span class="sxs-lookup"><span data-stu-id="9c730-744">String</span></span>|<span data-ttu-id="9c730-745">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="9c730-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c730-746">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-746">Requirements</span></span>

|<span data-ttu-id="9c730-747">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-747">Requirement</span></span>| <span data-ttu-id="9c730-748">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-749">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-749">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-750">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-750">1.0</span></span>|
|[<span data-ttu-id="9c730-751">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-752">ReadItem</span></span>|
|[<span data-ttu-id="9c730-753">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-754">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c730-755">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9c730-755">Returns:</span></span>

<span data-ttu-id="9c730-p146">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="9c730-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="9c730-758">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9c730-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="9c730-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9c730-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9c730-760">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9c730-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-761">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c730-p147">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="9c730-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9c730-765">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="9c730-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9c730-766">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9c730-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9c730-p148">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="9c730-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c730-769">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-769">Requirements</span></span>

|<span data-ttu-id="9c730-770">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-770">Requirement</span></span>| <span data-ttu-id="9c730-771">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-772">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-772">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-773">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-773">1.0</span></span>|
|[<span data-ttu-id="9c730-774">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-775">ReadItem</span></span>|
|[<span data-ttu-id="9c730-776">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-777">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c730-778">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9c730-778">Returns:</span></span>

<span data-ttu-id="9c730-p149">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="9c730-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9c730-781">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9c730-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9c730-782">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9c730-783">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-783">Example</span></span>

<span data-ttu-id="9c730-784">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="9c730-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9c730-785">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="9c730-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9c730-786">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9c730-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9c730-787">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="9c730-787">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c730-788">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="9c730-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9c730-p150">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="9c730-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-791">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-791">Parameters:</span></span>

|<span data-ttu-id="9c730-792">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-792">Name</span></span>| <span data-ttu-id="9c730-793">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-793">Type</span></span>| <span data-ttu-id="9c730-794">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9c730-795">String</span><span class="sxs-lookup"><span data-stu-id="9c730-795">String</span></span>|<span data-ttu-id="9c730-796">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="9c730-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c730-797">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-797">Requirements</span></span>

|<span data-ttu-id="9c730-798">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-798">Requirement</span></span>| <span data-ttu-id="9c730-799">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-800">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-800">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-801">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-801">1.0</span></span>|
|[<span data-ttu-id="9c730-802">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-803">ReadItem</span></span>|
|[<span data-ttu-id="9c730-804">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-805">Leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c730-806">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9c730-806">Returns:</span></span>

<span data-ttu-id="9c730-807">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9c730-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9c730-808">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="9c730-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9c730-809">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="9c730-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9c730-810">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9c730-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9c730-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9c730-812">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9c730-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9c730-p151">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="9c730-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-816">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-816">Parameters:</span></span>

|<span data-ttu-id="9c730-817">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-817">Name</span></span>| <span data-ttu-id="9c730-818">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-818">Type</span></span>| <span data-ttu-id="9c730-819">Atributos</span><span class="sxs-lookup"><span data-stu-id="9c730-819">Attributes</span></span>| <span data-ttu-id="9c730-820">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9c730-821">function</span><span class="sxs-lookup"><span data-stu-id="9c730-821">function</span></span>||<span data-ttu-id="9c730-822">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9c730-823">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c730-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9c730-824">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="9c730-824">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9c730-825">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-825">Object</span></span>| <span data-ttu-id="9c730-826">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-826">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-827">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-827">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9c730-828">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c730-829">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-829">Requirements</span></span>

|<span data-ttu-id="9c730-830">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-830">Requirement</span></span>| <span data-ttu-id="9c730-831">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-832">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-832">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-833">1.0</span><span class="sxs-lookup"><span data-stu-id="9c730-833">1.0</span></span>|
|[<span data-ttu-id="9c730-834">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c730-835">ReadItem</span></span>|
|[<span data-ttu-id="9c730-836">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-837">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9c730-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-838">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-838">Example</span></span>

<span data-ttu-id="9c730-p154">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9c730-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9c730-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9c730-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9c730-843">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9c730-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9c730-p155">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="9c730-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c730-848">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="9c730-848">Parameters:</span></span>

|<span data-ttu-id="9c730-849">Nome</span><span class="sxs-lookup"><span data-stu-id="9c730-849">Name</span></span>| <span data-ttu-id="9c730-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="9c730-850">Type</span></span>| <span data-ttu-id="9c730-851">Atributos</span><span class="sxs-lookup"><span data-stu-id="9c730-851">Attributes</span></span>| <span data-ttu-id="9c730-852">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9c730-853">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9c730-853">String</span></span>||<span data-ttu-id="9c730-p156">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9c730-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="9c730-856">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-856">Object</span></span>| <span data-ttu-id="9c730-857">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-857">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-858">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9c730-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9c730-859">Objeto</span><span class="sxs-lookup"><span data-stu-id="9c730-859">Object</span></span>| <span data-ttu-id="9c730-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-860">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-861">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9c730-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9c730-862">function</span><span class="sxs-lookup"><span data-stu-id="9c730-862">function</span></span>| <span data-ttu-id="9c730-863">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9c730-863">&lt;optional&gt;</span></span>|<span data-ttu-id="9c730-864">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c730-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9c730-865">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="9c730-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c730-866">Erros</span><span class="sxs-lookup"><span data-stu-id="9c730-866">Errors</span></span>

| <span data-ttu-id="9c730-867">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9c730-867">Error code</span></span> | <span data-ttu-id="9c730-868">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c730-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9c730-869">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="9c730-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c730-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c730-870">Requirements</span></span>

|<span data-ttu-id="9c730-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="9c730-871">Requirement</span></span>| <span data-ttu-id="9c730-872">Valor</span><span class="sxs-lookup"><span data-stu-id="9c730-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c730-873">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9c730-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c730-874">1.1</span><span class="sxs-lookup"><span data-stu-id="9c730-874">1.1</span></span>|
|[<span data-ttu-id="9c730-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9c730-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c730-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9c730-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="9c730-877">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9c730-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9c730-878">Composição</span><span class="sxs-lookup"><span data-stu-id="9c730-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9c730-879">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9c730-879">Example</span></span>

<span data-ttu-id="9c730-880">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="9c730-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```