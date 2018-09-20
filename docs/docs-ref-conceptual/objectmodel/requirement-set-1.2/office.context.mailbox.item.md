
# <a name="item"></a><span data-ttu-id="14a27-101">item</span><span class="sxs-lookup"><span data-stu-id="14a27-101">item</span></span>

### <span data-ttu-id="14a27-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="14a27-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="14a27-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="14a27-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-106">Requirements</span></span>

|<span data-ttu-id="14a27-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-107">Requirement</span></span>| <span data-ttu-id="14a27-108">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-110">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-110">1.0</span></span>|
|[<span data-ttu-id="14a27-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="14a27-112">Restricted</span></span>|
|[<span data-ttu-id="14a27-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="14a27-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-115">Example</span></span>

<span data-ttu-id="14a27-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="14a27-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="14a27-117">Membros</span><span class="sxs-lookup"><span data-stu-id="14a27-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="14a27-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="14a27-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="14a27-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-121">Certos tipos de arquivos são bloqueados pelo Outlook devido aos potenciais problemas de segurança e não, portanto, serão retornados.</span><span class="sxs-lookup"><span data-stu-id="14a27-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="14a27-122">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="14a27-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-123">Type:</span></span>

*   <span data-ttu-id="14a27-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="14a27-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-125">Requirements</span></span>

|<span data-ttu-id="14a27-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-126">Requirement</span></span>| <span data-ttu-id="14a27-127">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-128">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-129">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-129">1.0</span></span>|
|[<span data-ttu-id="14a27-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-131">ReadItem</span></span>|
|[<span data-ttu-id="14a27-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-133">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-134">Example</span></span>

<span data-ttu-id="14a27-135">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="14a27-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="14a27-137">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="14a27-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="14a27-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-139">Type:</span></span>

*   [<span data-ttu-id="14a27-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="14a27-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="14a27-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-141">Requirements</span></span>

|<span data-ttu-id="14a27-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-142">Requirement</span></span>| <span data-ttu-id="14a27-143">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-145">1.1</span><span class="sxs-lookup"><span data-stu-id="14a27-145">1.1</span></span>|
|[<span data-ttu-id="14a27-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-147">ReadItem</span></span>|
|[<span data-ttu-id="14a27-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-149">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="14a27-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="14a27-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="14a27-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="14a27-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-153">Type:</span></span>

*   [<span data-ttu-id="14a27-154">Corpo</span><span class="sxs-lookup"><span data-stu-id="14a27-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="14a27-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-155">Requirements</span></span>

|<span data-ttu-id="14a27-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-156">Requirement</span></span>| <span data-ttu-id="14a27-157">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-159">1.1</span><span class="sxs-lookup"><span data-stu-id="14a27-159">1.1</span></span>|
|[<span data-ttu-id="14a27-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-161">ReadItem</span></span>|
|[<span data-ttu-id="14a27-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="14a27-164">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="14a27-165">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="14a27-166">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-167">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-167">Read mode</span></span>

<span data-ttu-id="14a27-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="14a27-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-170">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-170">Compose mode</span></span>

<span data-ttu-id="14a27-171">O `cc` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-172">Type:</span></span>

*   <span data-ttu-id="14a27-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-174">Requirements</span></span>

|<span data-ttu-id="14a27-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-175">Requirement</span></span>| <span data-ttu-id="14a27-176">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-178">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-178">1.0</span></span>|
|[<span data-ttu-id="14a27-179">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-180">ReadItem</span></span>|
|[<span data-ttu-id="14a27-181">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-182">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="14a27-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="14a27-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="14a27-185">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="14a27-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="14a27-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="14a27-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="14a27-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="14a27-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-190">Type:</span></span>

*   <span data-ttu-id="14a27-191">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-192">Requirements</span></span>

|<span data-ttu-id="14a27-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-193">Requirement</span></span>| <span data-ttu-id="14a27-194">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-196">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-196">1.0</span></span>|
|[<span data-ttu-id="14a27-197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-198">ReadItem</span></span>|
|[<span data-ttu-id="14a27-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-200">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="14a27-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="14a27-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="14a27-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-204">Type:</span></span>

*   <span data-ttu-id="14a27-205">Data</span><span class="sxs-lookup"><span data-stu-id="14a27-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-206">Requirements</span></span>

|<span data-ttu-id="14a27-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-207">Requirement</span></span>| <span data-ttu-id="14a27-208">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-210">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-210">1.0</span></span>|
|[<span data-ttu-id="14a27-211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-212">ReadItem</span></span>|
|[<span data-ttu-id="14a27-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-214">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="14a27-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="14a27-216">dateTimeModified :Date</span></span>

<span data-ttu-id="14a27-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-219">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-220">Type:</span></span>

*   <span data-ttu-id="14a27-221">Data</span><span class="sxs-lookup"><span data-stu-id="14a27-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-222">Requirements</span></span>

|<span data-ttu-id="14a27-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-223">Requirement</span></span>| <span data-ttu-id="14a27-224">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-226">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-226">1.0</span></span>|
|[<span data-ttu-id="14a27-227">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-228">ReadItem</span></span>|
|[<span data-ttu-id="14a27-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-230">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-231">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="14a27-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="14a27-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="14a27-233">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="14a27-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="14a27-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="14a27-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-236">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-236">Read mode</span></span>

<span data-ttu-id="14a27-237">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="14a27-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-238">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-238">Compose mode</span></span>

<span data-ttu-id="14a27-239">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="14a27-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="14a27-240">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="14a27-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-241">Type:</span></span>

*   <span data-ttu-id="14a27-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="14a27-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-243">Requirements</span></span>

|<span data-ttu-id="14a27-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-244">Requirement</span></span>| <span data-ttu-id="14a27-245">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-246">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-247">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-247">1.0</span></span>|
|[<span data-ttu-id="14a27-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-249">ReadItem</span></span>|
|[<span data-ttu-id="14a27-250">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-251">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-252">Example</span></span>

<span data-ttu-id="14a27-253">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="14a27-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="14a27-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="14a27-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="14a27-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="14a27-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="14a27-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-259">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `from` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14a27-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-260">Type:</span></span>

*   [<span data-ttu-id="14a27-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14a27-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="14a27-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-262">Requirements</span></span>

|<span data-ttu-id="14a27-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-263">Requirement</span></span>| <span data-ttu-id="14a27-264">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-265">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-266">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-266">1.0</span></span>|
|[<span data-ttu-id="14a27-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-268">ReadItem</span></span>|
|[<span data-ttu-id="14a27-269">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-270">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="14a27-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="14a27-271">internetMessageId :String</span></span>

<span data-ttu-id="14a27-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-274">Type:</span></span>

*   <span data-ttu-id="14a27-275">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-276">Requirements</span></span>

|<span data-ttu-id="14a27-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-277">Requirement</span></span>| <span data-ttu-id="14a27-278">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-280">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-280">1.0</span></span>|
|[<span data-ttu-id="14a27-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-282">ReadItem</span></span>|
|[<span data-ttu-id="14a27-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-284">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="14a27-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="14a27-286">itemClass :String</span></span>

<span data-ttu-id="14a27-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="14a27-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="14a27-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-291">Type</span></span> | <span data-ttu-id="14a27-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-292">Description</span></span> | <span data-ttu-id="14a27-293">classe de item</span><span class="sxs-lookup"><span data-stu-id="14a27-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="14a27-294">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="14a27-294">Appointment items</span></span> | <span data-ttu-id="14a27-295">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="14a27-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="14a27-296">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="14a27-296">Message items</span></span> | <span data-ttu-id="14a27-297">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="14a27-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="14a27-298">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="14a27-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-299">Type:</span></span>

*   <span data-ttu-id="14a27-300">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-301">Requirements</span></span>

|<span data-ttu-id="14a27-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-302">Requirement</span></span>| <span data-ttu-id="14a27-303">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-305">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-305">1.0</span></span>|
|[<span data-ttu-id="14a27-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-307">ReadItem</span></span>|
|[<span data-ttu-id="14a27-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-309">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="14a27-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="14a27-311">(nullable) itemId :String</span></span>

<span data-ttu-id="14a27-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-314">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="14a27-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="14a27-315">O `itemId` propriedade não é idêntica à identificação de entrada do Outlook ou o ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="14a27-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="14a27-316">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do requisito definido 1.3.</span><span class="sxs-lookup"><span data-stu-id="14a27-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="14a27-317">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="14a27-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-318">Type:</span></span>

*   <span data-ttu-id="14a27-319">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-320">Requirements</span></span>

|<span data-ttu-id="14a27-321">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-321">Requirement</span></span>| <span data-ttu-id="14a27-322">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-323">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-324">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-324">1.0</span></span>|
|[<span data-ttu-id="14a27-325">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-326">ReadItem</span></span>|
|[<span data-ttu-id="14a27-327">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-328">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-329">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-329">Example</span></span>

<span data-ttu-id="14a27-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="14a27-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="14a27-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="14a27-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="14a27-333">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="14a27-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="14a27-334">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-335">Type:</span></span>

*   [<span data-ttu-id="14a27-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="14a27-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="14a27-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-337">Requirements</span></span>

|<span data-ttu-id="14a27-338">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-338">Requirement</span></span>| <span data-ttu-id="14a27-339">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-340">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-341">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-341">1.0</span></span>|
|[<span data-ttu-id="14a27-342">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-343">ReadItem</span></span>|
|[<span data-ttu-id="14a27-344">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-345">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="14a27-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="14a27-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="14a27-348">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-349">Read mode</span></span>

<span data-ttu-id="14a27-350">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-351">Compose mode</span></span>

<span data-ttu-id="14a27-352">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-353">Type:</span></span>

*   <span data-ttu-id="14a27-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="14a27-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-355">Requirements</span></span>

|<span data-ttu-id="14a27-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-356">Requirement</span></span>| <span data-ttu-id="14a27-357">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-358">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-359">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-359">1.0</span></span>|
|[<span data-ttu-id="14a27-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-361">ReadItem</span></span>|
|[<span data-ttu-id="14a27-362">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-363">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-364">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="14a27-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="14a27-365">normalizedSubject :String</span></span>

<span data-ttu-id="14a27-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="14a27-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="14a27-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-370">Type:</span></span>

*   <span data-ttu-id="14a27-371">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-372">Requirements</span></span>

|<span data-ttu-id="14a27-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-373">Requirement</span></span>| <span data-ttu-id="14a27-374">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-376">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-376">1.0</span></span>|
|[<span data-ttu-id="14a27-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-378">ReadItem</span></span>|
|[<span data-ttu-id="14a27-379">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-380">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-381">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="14a27-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="14a27-383">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="14a27-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="14a27-384">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-385">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-385">Read mode</span></span>

<span data-ttu-id="14a27-386">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="14a27-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-387">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-387">Compose mode</span></span>

<span data-ttu-id="14a27-388">O `optionalAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes opcionais para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="14a27-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-389">Type:</span></span>

*   <span data-ttu-id="14a27-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-391">Requirements</span></span>

|<span data-ttu-id="14a27-392">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-392">Requirement</span></span>| <span data-ttu-id="14a27-393">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-394">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-395">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-395">1.0</span></span>|
|[<span data-ttu-id="14a27-396">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-397">ReadItem</span></span>|
|[<span data-ttu-id="14a27-398">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-399">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-400">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="14a27-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="14a27-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="14a27-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-404">Type:</span></span>

*   [<span data-ttu-id="14a27-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14a27-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="14a27-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-406">Requirements</span></span>

|<span data-ttu-id="14a27-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-407">Requirement</span></span>| <span data-ttu-id="14a27-408">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-410">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-410">1.0</span></span>|
|[<span data-ttu-id="14a27-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-412">ReadItem</span></span>|
|[<span data-ttu-id="14a27-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-414">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="14a27-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="14a27-417">Fornece acesso aos participantes necessários de um evento.</span><span class="sxs-lookup"><span data-stu-id="14a27-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="14a27-418">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-419">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-419">Read mode</span></span>

<span data-ttu-id="14a27-420">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="14a27-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-421">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-421">Compose mode</span></span>

<span data-ttu-id="14a27-422">O `requiredAttendees` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os participantes necessários para uma reunião.</span><span class="sxs-lookup"><span data-stu-id="14a27-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-423">Type:</span></span>

*   <span data-ttu-id="14a27-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-425">Requirements</span></span>

|<span data-ttu-id="14a27-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-426">Requirement</span></span>| <span data-ttu-id="14a27-427">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-429">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-429">1.0</span></span>|
|[<span data-ttu-id="14a27-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-431">ReadItem</span></span>|
|[<span data-ttu-id="14a27-432">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-433">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="14a27-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="14a27-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="14a27-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="14a27-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="14a27-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="14a27-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-440">O `recipientType` propriedade do `EmailAddressDetails` do objeto no `sender` propriedadeé `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14a27-440">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-441">Type:</span></span>

*   [<span data-ttu-id="14a27-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14a27-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="14a27-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-443">Requirements</span></span>

|<span data-ttu-id="14a27-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-444">Requirement</span></span>| <span data-ttu-id="14a27-445">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-447">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-447">1.0</span></span>|
|[<span data-ttu-id="14a27-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-449">ReadItem</span></span>|
|[<span data-ttu-id="14a27-450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-451">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="14a27-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="14a27-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="14a27-454">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="14a27-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="14a27-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="14a27-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-457">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-457">Read mode</span></span>

<span data-ttu-id="14a27-458">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="14a27-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-459">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-459">Compose mode</span></span>

<span data-ttu-id="14a27-460">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="14a27-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="14a27-461">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="14a27-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-462">Type:</span></span>

*   <span data-ttu-id="14a27-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="14a27-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-464">Requirements</span></span>

|<span data-ttu-id="14a27-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-465">Requirement</span></span>| <span data-ttu-id="14a27-466">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-468">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-468">1.0</span></span>|
|[<span data-ttu-id="14a27-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-470">ReadItem</span></span>|
|[<span data-ttu-id="14a27-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-472">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-473">Example</span></span>

<span data-ttu-id="14a27-474">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="14a27-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="14a27-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="14a27-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="14a27-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="14a27-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="14a27-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="14a27-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-478">Read mode</span></span>

<span data-ttu-id="14a27-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="14a27-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="14a27-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-481">Compose mode</span></span>

<span data-ttu-id="14a27-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="14a27-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14a27-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-483">Type:</span></span>

*   <span data-ttu-id="14a27-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="14a27-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-485">Requirements</span></span>

|<span data-ttu-id="14a27-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-486">Requirement</span></span>| <span data-ttu-id="14a27-487">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-489">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-489">1.0</span></span>|
|[<span data-ttu-id="14a27-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-491">ReadItem</span></span>|
|[<span data-ttu-id="14a27-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-493">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="14a27-494">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="14a27-495">Fornece acesso aos destinatários na linha **para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="14a27-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14a27-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-497">Read mode</span></span>

<span data-ttu-id="14a27-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="14a27-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="14a27-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="14a27-500">Compose mode</span></span>

<span data-ttu-id="14a27-501">O `to` propriedade retornará uma `Recipients` objeto que fornece os métodos para obter ou atualizar os destinatários na linha **para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="14a27-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="14a27-502">Type:</span></span>

*   <span data-ttu-id="14a27-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="14a27-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-504">Requirements</span></span>

|<span data-ttu-id="14a27-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-505">Requirement</span></span>| <span data-ttu-id="14a27-506">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-508">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-508">1.0</span></span>|
|[<span data-ttu-id="14a27-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-510">ReadItem</span></span>|
|[<span data-ttu-id="14a27-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-512">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="14a27-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="14a27-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="14a27-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14a27-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14a27-516">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="14a27-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="14a27-517">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="14a27-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="14a27-518">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="14a27-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-519">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-519">Parameters:</span></span>

|<span data-ttu-id="14a27-520">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-520">Name</span></span>| <span data-ttu-id="14a27-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-521">Type</span></span>| <span data-ttu-id="14a27-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-522">Attributes</span></span>| <span data-ttu-id="14a27-523">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="14a27-524">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-524">String</span></span>||<span data-ttu-id="14a27-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14a27-527">String</span><span class="sxs-lookup"><span data-stu-id="14a27-527">String</span></span>||<span data-ttu-id="14a27-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14a27-530">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-530">Object</span></span>| <span data-ttu-id="14a27-531">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-531">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-532">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14a27-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-533">Object</span></span>| <span data-ttu-id="14a27-534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-534">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-535">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14a27-536">function</span><span class="sxs-lookup"><span data-stu-id="14a27-536">function</span></span>| <span data-ttu-id="14a27-537">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-537">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-538">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14a27-539">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14a27-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14a27-540">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="14a27-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14a27-541">Erros</span><span class="sxs-lookup"><span data-stu-id="14a27-541">Errors</span></span>

| <span data-ttu-id="14a27-542">Código de erro</span><span class="sxs-lookup"><span data-stu-id="14a27-542">Error code</span></span> | <span data-ttu-id="14a27-543">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="14a27-544">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="14a27-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="14a27-545">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="14a27-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14a27-546">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="14a27-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-547">Requirements</span></span>

|<span data-ttu-id="14a27-548">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-548">Requirement</span></span>| <span data-ttu-id="14a27-549">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-550">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-551">1.1</span><span class="sxs-lookup"><span data-stu-id="14a27-551">1.1</span></span>|
|[<span data-ttu-id="14a27-552">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14a27-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="14a27-554">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-555">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="14a27-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14a27-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14a27-558">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="14a27-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="14a27-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="14a27-562">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="14a27-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="14a27-563">Se o suplemento do Office estiver sendo executado no Outlook Web App, o `addItemAttachmentAsync` método pode anexar itens a itens que não sejam o item que você está editando; No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="14a27-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-564">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-564">Parameters:</span></span>

|<span data-ttu-id="14a27-565">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-565">Name</span></span>| <span data-ttu-id="14a27-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-566">Type</span></span>| <span data-ttu-id="14a27-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-567">Attributes</span></span>| <span data-ttu-id="14a27-568">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="14a27-569">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-569">String</span></span>||<span data-ttu-id="14a27-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14a27-572">String</span><span class="sxs-lookup"><span data-stu-id="14a27-572">String</span></span>||<span data-ttu-id="14a27-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14a27-575">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-575">Object</span></span>| <span data-ttu-id="14a27-576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-576">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-577">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14a27-578">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-578">Object</span></span>| <span data-ttu-id="14a27-579">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-579">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-580">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14a27-581">function</span><span class="sxs-lookup"><span data-stu-id="14a27-581">function</span></span>| <span data-ttu-id="14a27-582">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-582">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-583">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14a27-584">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14a27-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14a27-585">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="14a27-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14a27-586">Erros</span><span class="sxs-lookup"><span data-stu-id="14a27-586">Errors</span></span>

| <span data-ttu-id="14a27-587">Código de erro</span><span class="sxs-lookup"><span data-stu-id="14a27-587">Error code</span></span> | <span data-ttu-id="14a27-588">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14a27-589">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="14a27-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-590">Requirements</span></span>

|<span data-ttu-id="14a27-591">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-591">Requirement</span></span>| <span data-ttu-id="14a27-592">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-593">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-594">1.1</span><span class="sxs-lookup"><span data-stu-id="14a27-594">1.1</span></span>|
|[<span data-ttu-id="14a27-595">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14a27-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="14a27-597">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-598">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-599">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-599">Example</span></span>

<span data-ttu-id="14a27-600">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="14a27-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="14a27-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="14a27-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="14a27-602">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="14a27-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-603">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14a27-604">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="14a27-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14a27-605">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="14a27-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="14a27-p137">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="14a27-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-609">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-609">Parameters:</span></span>

|<span data-ttu-id="14a27-610">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-610">Name</span></span>| <span data-ttu-id="14a27-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-611">Type</span></span>| <span data-ttu-id="14a27-612">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14a27-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14a27-613">String &#124; Object</span></span>| |<span data-ttu-id="14a27-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="14a27-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14a27-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="14a27-616">**OR**</span></span><br/><span data-ttu-id="14a27-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14a27-619">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-619">String</span></span> | <span data-ttu-id="14a27-620">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-620">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="14a27-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="14a27-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="14a27-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-624">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-625">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="14a27-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="14a27-626">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-626">String</span></span> | | <span data-ttu-id="14a27-p141">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="14a27-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="14a27-629">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-629">String</span></span> | | <span data-ttu-id="14a27-630">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="14a27-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="14a27-631">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-631">String</span></span> | | <span data-ttu-id="14a27-p142">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="14a27-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="14a27-634">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-634">String</span></span> | | <span data-ttu-id="14a27-p143">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="14a27-638">function</span><span class="sxs-lookup"><span data-stu-id="14a27-638">function</span></span> | <span data-ttu-id="14a27-639">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-639">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-640">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-641">Requirements</span></span>

|<span data-ttu-id="14a27-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-642">Requirement</span></span>| <span data-ttu-id="14a27-643">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-644">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-645">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-645">1.0</span></span>|
|[<span data-ttu-id="14a27-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-647">ReadItem</span></span>|
|[<span data-ttu-id="14a27-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-649">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14a27-650">Exemplos</span><span class="sxs-lookup"><span data-stu-id="14a27-650">Examples</span></span>

<span data-ttu-id="14a27-651">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="14a27-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="14a27-652">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="14a27-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="14a27-653">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="14a27-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14a27-654">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="14a27-654">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="14a27-655">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="14a27-655">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="14a27-656">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="14a27-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="14a27-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="14a27-658">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="14a27-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-659">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-659">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14a27-660">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="14a27-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14a27-661">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="14a27-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="14a27-p144">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="14a27-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-665">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-665">Parameters:</span></span>

|<span data-ttu-id="14a27-666">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-666">Name</span></span>| <span data-ttu-id="14a27-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-667">Type</span></span>| <span data-ttu-id="14a27-668">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14a27-669">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14a27-669">String &#124; Object</span></span>| | <span data-ttu-id="14a27-p145">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="14a27-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14a27-672">**OU**</span><span class="sxs-lookup"><span data-stu-id="14a27-672">**OR**</span></span><br/><span data-ttu-id="14a27-p146">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14a27-675">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-675">String</span></span> | <span data-ttu-id="14a27-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-676">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="14a27-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="14a27-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="14a27-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-680">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-681">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="14a27-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="14a27-682">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-682">String</span></span> | | <span data-ttu-id="14a27-p148">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="14a27-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="14a27-685">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-685">String</span></span> | | <span data-ttu-id="14a27-686">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="14a27-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="14a27-687">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-687">String</span></span> | | <span data-ttu-id="14a27-p149">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="14a27-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="14a27-690">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-690">String</span></span> | | <span data-ttu-id="14a27-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="14a27-694">function</span><span class="sxs-lookup"><span data-stu-id="14a27-694">function</span></span> | <span data-ttu-id="14a27-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-695">&lt;optional&gt;</span></span> | <span data-ttu-id="14a27-696">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-697">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-697">Requirements</span></span>

|<span data-ttu-id="14a27-698">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-698">Requirement</span></span>| <span data-ttu-id="14a27-699">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-700">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-701">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-701">1.0</span></span>|
|[<span data-ttu-id="14a27-702">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-703">ReadItem</span></span>|
|[<span data-ttu-id="14a27-704">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-705">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14a27-706">Exemplos</span><span class="sxs-lookup"><span data-stu-id="14a27-706">Examples</span></span>

<span data-ttu-id="14a27-707">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="14a27-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="14a27-708">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="14a27-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="14a27-709">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="14a27-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14a27-710">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="14a27-710">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="14a27-711">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="14a27-711">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="14a27-712">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="14a27-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="14a27-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="14a27-714">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="14a27-714">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-715">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-716">Requirements</span></span>

|<span data-ttu-id="14a27-717">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-717">Requirement</span></span>| <span data-ttu-id="14a27-718">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-719">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-720">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-720">1.0</span></span>|
|[<span data-ttu-id="14a27-721">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-722">ReadItem</span></span>|
|[<span data-ttu-id="14a27-723">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-724">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-725">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-725">Returns:</span></span>

<span data-ttu-id="14a27-726">Tipo: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="14a27-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="14a27-727">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-727">Example</span></span>

<span data-ttu-id="14a27-728">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-728">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="14a27-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="14a27-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="14a27-730">Obtém uma matriz de todas as entidades do tipo de entidade especificada encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="14a27-730">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-731">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-731">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-732">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-732">Parameters:</span></span>

|<span data-ttu-id="14a27-733">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-733">Name</span></span>| <span data-ttu-id="14a27-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-734">Type</span></span>| <span data-ttu-id="14a27-735">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="14a27-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="14a27-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="14a27-737">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="14a27-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14a27-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-738">Requirements</span></span>

|<span data-ttu-id="14a27-739">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-739">Requirement</span></span>| <span data-ttu-id="14a27-740">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-741">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-742">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-742">1.0</span></span>|
|[<span data-ttu-id="14a27-743">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-744">Restrito</span><span class="sxs-lookup"><span data-stu-id="14a27-744">Restricted</span></span>|
|[<span data-ttu-id="14a27-745">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-746">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-747">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-747">Returns:</span></span>

<span data-ttu-id="14a27-748">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="14a27-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="14a27-749">Se não há entidades do tipo especificado estiverem presentes no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="14a27-749">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="14a27-750">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="14a27-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="14a27-751">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="14a27-752">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="14a27-752">Value of `entityType`</span></span> | <span data-ttu-id="14a27-753">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="14a27-753">Type of objects in returned array</span></span> | <span data-ttu-id="14a27-754">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="14a27-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="14a27-755">String</span><span class="sxs-lookup"><span data-stu-id="14a27-755">String</span></span> | <span data-ttu-id="14a27-756">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="14a27-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="14a27-757">Contato</span><span class="sxs-lookup"><span data-stu-id="14a27-757">Contact</span></span> | <span data-ttu-id="14a27-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14a27-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="14a27-759">String</span><span class="sxs-lookup"><span data-stu-id="14a27-759">String</span></span> | <span data-ttu-id="14a27-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14a27-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="14a27-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="14a27-761">MeetingSuggestion</span></span> | <span data-ttu-id="14a27-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14a27-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="14a27-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="14a27-763">PhoneNumber</span></span> | <span data-ttu-id="14a27-764">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="14a27-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="14a27-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="14a27-765">TaskSuggestion</span></span> | <span data-ttu-id="14a27-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14a27-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="14a27-767">String</span><span class="sxs-lookup"><span data-stu-id="14a27-767">String</span></span> | <span data-ttu-id="14a27-768">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="14a27-768">**Restricted**</span></span> |

<span data-ttu-id="14a27-769">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="14a27-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="14a27-770">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-770">Example</span></span>

<span data-ttu-id="14a27-771">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representam endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="14a27-771">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="14a27-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="14a27-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="14a27-773">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="14a27-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-774">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-774">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14a27-775">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="14a27-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-776">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-776">Parameters:</span></span>

|<span data-ttu-id="14a27-777">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-777">Name</span></span>| <span data-ttu-id="14a27-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-778">Type</span></span>| <span data-ttu-id="14a27-779">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14a27-780">String</span><span class="sxs-lookup"><span data-stu-id="14a27-780">String</span></span>|<span data-ttu-id="14a27-781">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="14a27-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14a27-782">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-782">Requirements</span></span>

|<span data-ttu-id="14a27-783">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-783">Requirement</span></span>| <span data-ttu-id="14a27-784">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-785">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-786">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-786">1.0</span></span>|
|[<span data-ttu-id="14a27-787">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-788">ReadItem</span></span>|
|[<span data-ttu-id="14a27-789">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-790">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-791">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-791">Returns:</span></span>

<span data-ttu-id="14a27-p152">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="14a27-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="14a27-794">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="14a27-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="14a27-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="14a27-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="14a27-796">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="14a27-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-797">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-797">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14a27-p153">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="14a27-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="14a27-801">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="14a27-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="14a27-802">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="14a27-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="14a27-p154">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="14a27-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14a27-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-805">Requirements</span></span>

|<span data-ttu-id="14a27-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-806">Requirement</span></span>| <span data-ttu-id="14a27-807">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-809">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-809">1.0</span></span>|
|[<span data-ttu-id="14a27-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-811">ReadItem</span></span>|
|[<span data-ttu-id="14a27-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-813">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-814">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-814">Returns:</span></span>

<span data-ttu-id="14a27-p155">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="14a27-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="14a27-817">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="14a27-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="14a27-818">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="14a27-819">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-819">Example</span></span>

<span data-ttu-id="14a27-820">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="14a27-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="14a27-821">→ getRegExMatchesByName(name) (anulável) {matriz. < cadeia de caracteres >}</span><span class="sxs-lookup"><span data-stu-id="14a27-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="14a27-822">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="14a27-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14a27-823">Esse método não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="14a27-823">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14a27-824">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="14a27-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="14a27-p156">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="14a27-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-827">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-827">Parameters:</span></span>

|<span data-ttu-id="14a27-828">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-828">Name</span></span>| <span data-ttu-id="14a27-829">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-829">Type</span></span>| <span data-ttu-id="14a27-830">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14a27-831">String</span><span class="sxs-lookup"><span data-stu-id="14a27-831">String</span></span>|<span data-ttu-id="14a27-832">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="14a27-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14a27-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-833">Requirements</span></span>

|<span data-ttu-id="14a27-834">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-834">Requirement</span></span>| <span data-ttu-id="14a27-835">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-836">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-837">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-837">1.0</span></span>|
|[<span data-ttu-id="14a27-838">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-839">ReadItem</span></span>|
|[<span data-ttu-id="14a27-840">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-841">Leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-842">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-842">Returns:</span></span>

<span data-ttu-id="14a27-843">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="14a27-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="14a27-844">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="14a27-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="14a27-845">Matriz. < cadeia de caracteres ></span><span class="sxs-lookup"><span data-stu-id="14a27-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="14a27-846">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="14a27-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="14a27-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="14a27-848">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="14a27-p157">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="14a27-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-851">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-851">Parameters:</span></span>

|<span data-ttu-id="14a27-852">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-852">Name</span></span>| <span data-ttu-id="14a27-853">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-853">Type</span></span>| <span data-ttu-id="14a27-854">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-854">Attributes</span></span>| <span data-ttu-id="14a27-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="14a27-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="14a27-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="14a27-p158">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="14a27-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="14a27-860">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-860">Object</span></span>| <span data-ttu-id="14a27-861">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-861">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-862">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14a27-863">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-863">Object</span></span>| <span data-ttu-id="14a27-864">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-864">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-865">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14a27-866">function</span><span class="sxs-lookup"><span data-stu-id="14a27-866">function</span></span>||<span data-ttu-id="14a27-867">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14a27-868">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="14a27-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="14a27-869">Para acessar a propriedade source que a seleção é proveniente, chamar `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="14a27-869">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14a27-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-870">Requirements</span></span>

|<span data-ttu-id="14a27-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-871">Requirement</span></span>| <span data-ttu-id="14a27-872">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-873">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-874">1.2</span><span class="sxs-lookup"><span data-stu-id="14a27-874">1.2</span></span>|
|[<span data-ttu-id="14a27-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14a27-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="14a27-877">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-878">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="14a27-879">Retorna:</span><span class="sxs-lookup"><span data-stu-id="14a27-879">Returns:</span></span>

<span data-ttu-id="14a27-880">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="14a27-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="14a27-881">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="14a27-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="14a27-882">String</span><span class="sxs-lookup"><span data-stu-id="14a27-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="14a27-883">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-883">Example</span></span>

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="14a27-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14a27-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="14a27-885">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="14a27-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="14a27-p160">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="14a27-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-889">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-889">Parameters:</span></span>

|<span data-ttu-id="14a27-890">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-890">Name</span></span>| <span data-ttu-id="14a27-891">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-891">Type</span></span>| <span data-ttu-id="14a27-892">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-892">Attributes</span></span>| <span data-ttu-id="14a27-893">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="14a27-894">function</span><span class="sxs-lookup"><span data-stu-id="14a27-894">function</span></span>||<span data-ttu-id="14a27-895">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14a27-896">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14a27-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="14a27-897">Este objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações para a propriedade personalizada definida de volta ao servidor.</span><span class="sxs-lookup"><span data-stu-id="14a27-897">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="14a27-898">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-898">Object</span></span>| <span data-ttu-id="14a27-899">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-899">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-900">Os desenvolvedores podem fornecer qualquer objeto que desejam acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-900">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="14a27-901">Este objeto pode ser acessado pelo `asyncResult.asyncContext` propriedade na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14a27-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-902">Requirements</span></span>

|<span data-ttu-id="14a27-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-903">Requirement</span></span>| <span data-ttu-id="14a27-904">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-906">1.0</span><span class="sxs-lookup"><span data-stu-id="14a27-906">1.0</span></span>|
|[<span data-ttu-id="14a27-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14a27-908">ReadItem</span></span>|
|[<span data-ttu-id="14a27-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-910">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="14a27-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-911">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-911">Example</span></span>

<span data-ttu-id="14a27-p163">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="14a27-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="14a27-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14a27-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="14a27-916">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="14a27-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="14a27-p164">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="14a27-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-921">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-921">Parameters:</span></span>

|<span data-ttu-id="14a27-922">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-922">Name</span></span>| <span data-ttu-id="14a27-923">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-923">Type</span></span>| <span data-ttu-id="14a27-924">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-924">Attributes</span></span>| <span data-ttu-id="14a27-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="14a27-926">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="14a27-926">String</span></span>||<span data-ttu-id="14a27-p165">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14a27-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="14a27-929">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-929">Object</span></span>| <span data-ttu-id="14a27-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-930">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-931">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14a27-932">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-932">Object</span></span>| <span data-ttu-id="14a27-933">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-933">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-934">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14a27-935">function</span><span class="sxs-lookup"><span data-stu-id="14a27-935">function</span></span>| <span data-ttu-id="14a27-936">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-936">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-937">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14a27-938">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="14a27-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14a27-939">Erros</span><span class="sxs-lookup"><span data-stu-id="14a27-939">Errors</span></span>

| <span data-ttu-id="14a27-940">Código de erro</span><span class="sxs-lookup"><span data-stu-id="14a27-940">Error code</span></span> | <span data-ttu-id="14a27-941">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="14a27-942">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="14a27-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-943">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-943">Requirements</span></span>

|<span data-ttu-id="14a27-944">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-944">Requirement</span></span>| <span data-ttu-id="14a27-945">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-946">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-947">1.1</span><span class="sxs-lookup"><span data-stu-id="14a27-947">1.1</span></span>|
|[<span data-ttu-id="14a27-948">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14a27-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="14a27-950">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-951">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-952">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-952">Example</span></span>

<span data-ttu-id="14a27-953">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="14a27-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="14a27-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="14a27-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="14a27-955">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="14a27-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="14a27-p166">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="14a27-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14a27-959">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="14a27-959">Parameters:</span></span>

|<span data-ttu-id="14a27-960">Nome</span><span class="sxs-lookup"><span data-stu-id="14a27-960">Name</span></span>| <span data-ttu-id="14a27-961">Tipo</span><span class="sxs-lookup"><span data-stu-id="14a27-961">Type</span></span>| <span data-ttu-id="14a27-962">Atributos</span><span class="sxs-lookup"><span data-stu-id="14a27-962">Attributes</span></span>| <span data-ttu-id="14a27-963">Descrição</span><span class="sxs-lookup"><span data-stu-id="14a27-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="14a27-964">String</span><span class="sxs-lookup"><span data-stu-id="14a27-964">String</span></span>||<span data-ttu-id="14a27-p167">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="14a27-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="14a27-968">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-968">Object</span></span>| <span data-ttu-id="14a27-969">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-969">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-970">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="14a27-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14a27-971">Objeto</span><span class="sxs-lookup"><span data-stu-id="14a27-971">Object</span></span>| <span data-ttu-id="14a27-972">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-972">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-973">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="14a27-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="14a27-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="14a27-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="14a27-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="14a27-975">&lt;optional&gt;</span></span>|<span data-ttu-id="14a27-p168">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="14a27-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="14a27-p169">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="14a27-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="14a27-980">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="14a27-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="14a27-981">function</span><span class="sxs-lookup"><span data-stu-id="14a27-981">function</span></span>||<span data-ttu-id="14a27-982">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14a27-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14a27-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="14a27-983">Requirements</span></span>

|<span data-ttu-id="14a27-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="14a27-984">Requirement</span></span>| <span data-ttu-id="14a27-985">Valor</span><span class="sxs-lookup"><span data-stu-id="14a27-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="14a27-986">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="14a27-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14a27-987">1.2</span><span class="sxs-lookup"><span data-stu-id="14a27-987">1.2</span></span>|
|[<span data-ttu-id="14a27-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="14a27-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14a27-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14a27-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="14a27-990">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="14a27-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14a27-991">Composição</span><span class="sxs-lookup"><span data-stu-id="14a27-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14a27-992">Exemplo</span><span class="sxs-lookup"><span data-stu-id="14a27-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```