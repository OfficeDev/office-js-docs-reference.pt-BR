
# <a name="context"></a><span data-ttu-id="e5994-101">context</span><span class="sxs-lookup"><span data-stu-id="e5994-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="e5994-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="e5994-102">[Office](Office.md).context</span></span>

<span data-ttu-id="e5994-p101">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="e5994-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5994-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5994-105">Requirements</span></span>

|<span data-ttu-id="e5994-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5994-106">Requirement</span></span>| <span data-ttu-id="e5994-107">Valor</span><span class="sxs-lookup"><span data-stu-id="e5994-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5994-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5994-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5994-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e5994-109">1.0</span></span>|
|[<span data-ttu-id="e5994-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5994-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5994-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e5994-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e5994-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e5994-112">Members and methods</span></span>

| <span data-ttu-id="e5994-113">Membro</span><span class="sxs-lookup"><span data-stu-id="e5994-113">Member</span></span> | <span data-ttu-id="e5994-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5994-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e5994-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e5994-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e5994-116">Membro</span><span class="sxs-lookup"><span data-stu-id="e5994-116">Member</span></span> |
| [<span data-ttu-id="e5994-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="e5994-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="e5994-118">Membro</span><span class="sxs-lookup"><span data-stu-id="e5994-118">Member</span></span> |
| [<span data-ttu-id="e5994-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e5994-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings) | <span data-ttu-id="e5994-120">Membro</span><span class="sxs-lookup"><span data-stu-id="e5994-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e5994-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="e5994-121">Namespaces</span></span>

<span data-ttu-id="e5994-122">[caixa de correio](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e Microsoft Outlook na web.</span><span class="sxs-lookup"><span data-stu-id="e5994-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="e5994-123">Membros</span><span class="sxs-lookup"><span data-stu-id="e5994-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="e5994-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="e5994-124">displayLanguage :String</span></span>

<span data-ttu-id="e5994-125">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e5994-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="e5994-126">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e5994-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="e5994-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5994-127">Type:</span></span>

*   <span data-ttu-id="e5994-128">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5994-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5994-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5994-129">Requirements</span></span>

|<span data-ttu-id="e5994-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5994-130">Requirement</span></span>| <span data-ttu-id="e5994-131">Valor</span><span class="sxs-lookup"><span data-stu-id="e5994-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5994-132">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5994-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5994-133">1.0</span><span class="sxs-lookup"><span data-stu-id="e5994-133">1.0</span></span>|
|[<span data-ttu-id="e5994-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5994-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5994-135">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e5994-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5994-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5994-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="e5994-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="e5994-137">officeTheme :Object</span></span>

<span data-ttu-id="e5994-138">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="e5994-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="e5994-139">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5994-139">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5994-p102">Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em \*\*Arquivo > Conta do Office > Tema da interface de usuário do Office \*\*, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e5994-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e5994-142">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5994-142">Type:</span></span>

*   <span data-ttu-id="e5994-143">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5994-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="e5994-144">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e5994-144">Properties:</span></span>

|<span data-ttu-id="e5994-145">Nome</span><span class="sxs-lookup"><span data-stu-id="e5994-145">Name</span></span>| <span data-ttu-id="e5994-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5994-146">Type</span></span>| <span data-ttu-id="e5994-147">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5994-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="e5994-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5994-148">String</span></span>|<span data-ttu-id="e5994-149">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="e5994-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="e5994-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5994-150">String</span></span>|<span data-ttu-id="e5994-151">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="e5994-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="e5994-152">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5994-152">String</span></span>|<span data-ttu-id="e5994-153">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="e5994-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="e5994-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5994-154">String</span></span>|<span data-ttu-id="e5994-155">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="e5994-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5994-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5994-156">Requirements</span></span>

|<span data-ttu-id="e5994-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5994-157">Requirement</span></span>| <span data-ttu-id="e5994-158">Valor</span><span class="sxs-lookup"><span data-stu-id="e5994-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5994-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5994-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5994-160">1.3</span><span class="sxs-lookup"><span data-stu-id="e5994-160">1.3</span></span>|
|[<span data-ttu-id="e5994-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5994-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5994-162">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e5994-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5994-163">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5994-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="e5994-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="e5994-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="e5994-165">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e5994-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e5994-166">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="e5994-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e5994-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5994-167">Type:</span></span>

*   [<span data-ttu-id="e5994-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e5994-168">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e5994-169">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5994-169">Requirements</span></span>

|<span data-ttu-id="e5994-170">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5994-170">Requirement</span></span>| <span data-ttu-id="e5994-171">Valor</span><span class="sxs-lookup"><span data-stu-id="e5994-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5994-172">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5994-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5994-173">1.0</span><span class="sxs-lookup"><span data-stu-id="e5994-173">1.0</span></span>|
|[<span data-ttu-id="e5994-174">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5994-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5994-175">Restrito</span><span class="sxs-lookup"><span data-stu-id="e5994-175">Restricted</span></span>|
|[<span data-ttu-id="e5994-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5994-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5994-177">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e5994-177">Compose or read</span></span>|