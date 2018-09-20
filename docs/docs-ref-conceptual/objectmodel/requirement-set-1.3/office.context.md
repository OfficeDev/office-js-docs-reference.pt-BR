
# <a name="context"></a><span data-ttu-id="0f0f1-101">context</span><span class="sxs-lookup"><span data-stu-id="0f0f1-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="0f0f1-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="0f0f1-102">[Office](Office.md).context</span></span>

<span data-ttu-id="0f0f1-p101">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="0f0f1-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f0f1-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0f0f1-105">Requirements</span></span>

|<span data-ttu-id="0f0f1-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="0f0f1-106">Requirement</span></span>| <span data-ttu-id="0f0f1-107">Valor</span><span class="sxs-lookup"><span data-stu-id="0f0f1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f0f1-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0f0f1-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f0f1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0f0f1-109">1.0</span></span>|
|[<span data-ttu-id="0f0f1-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0f0f1-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f0f1-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="0f0f1-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="0f0f1-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="0f0f1-112">Namespaces</span></span>

<span data-ttu-id="0f0f1-113">[caixa de correio](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e Microsoft Outlook na web.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="0f0f1-114">Membros</span><span class="sxs-lookup"><span data-stu-id="0f0f1-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="0f0f1-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="0f0f1-115">displayLanguage :String</span></span>

<span data-ttu-id="0f0f1-116">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="0f0f1-117">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="0f0f1-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="0f0f1-118">Type:</span></span>

*   <span data-ttu-id="0f0f1-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="0f0f1-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f0f1-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0f0f1-120">Requirements</span></span>

|<span data-ttu-id="0f0f1-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="0f0f1-121">Requirement</span></span>| <span data-ttu-id="0f0f1-122">Valor</span><span class="sxs-lookup"><span data-stu-id="0f0f1-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f0f1-123">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0f0f1-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f0f1-124">1.0</span><span class="sxs-lookup"><span data-stu-id="0f0f1-124">1.0</span></span>|
|[<span data-ttu-id="0f0f1-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0f0f1-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f0f1-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="0f0f1-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f0f1-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0f0f1-127">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="0f0f1-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="0f0f1-128">officeTheme :Object</span></span>

<span data-ttu-id="0f0f1-129">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="0f0f1-130">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-130">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f0f1-p102">Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em \*\*Arquivo > Conta do Office > Tema da interface de usuário do Office \*\*, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0f0f1-133">Tipo:</span><span class="sxs-lookup"><span data-stu-id="0f0f1-133">Type:</span></span>

*   <span data-ttu-id="0f0f1-134">Objeto</span><span class="sxs-lookup"><span data-stu-id="0f0f1-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="0f0f1-135">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="0f0f1-135">Properties:</span></span>

|<span data-ttu-id="0f0f1-136">Nome</span><span class="sxs-lookup"><span data-stu-id="0f0f1-136">Name</span></span>| <span data-ttu-id="0f0f1-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="0f0f1-137">Type</span></span>| <span data-ttu-id="0f0f1-138">Descrição</span><span class="sxs-lookup"><span data-stu-id="0f0f1-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="0f0f1-139">String</span><span class="sxs-lookup"><span data-stu-id="0f0f1-139">String</span></span>|<span data-ttu-id="0f0f1-140">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="0f0f1-141">String</span><span class="sxs-lookup"><span data-stu-id="0f0f1-141">String</span></span>|<span data-ttu-id="0f0f1-142">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="0f0f1-143">String</span><span class="sxs-lookup"><span data-stu-id="0f0f1-143">String</span></span>|<span data-ttu-id="0f0f1-144">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="0f0f1-145">String</span><span class="sxs-lookup"><span data-stu-id="0f0f1-145">String</span></span>|<span data-ttu-id="0f0f1-146">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f0f1-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0f0f1-147">Requirements</span></span>

|<span data-ttu-id="0f0f1-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="0f0f1-148">Requirement</span></span>| <span data-ttu-id="0f0f1-149">Valor</span><span class="sxs-lookup"><span data-stu-id="0f0f1-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f0f1-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0f0f1-150">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f0f1-151">1.3</span><span class="sxs-lookup"><span data-stu-id="0f0f1-151">1.3</span></span>|
|[<span data-ttu-id="0f0f1-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0f0f1-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f0f1-153">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="0f0f1-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f0f1-154">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0f0f1-154">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="0f0f1-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="0f0f1-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="0f0f1-156">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="0f0f1-157">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="0f0f1-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="0f0f1-158">Tipo:</span><span class="sxs-lookup"><span data-stu-id="0f0f1-158">Type:</span></span>

*   [<span data-ttu-id="0f0f1-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0f0f1-159">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="0f0f1-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0f0f1-160">Requirements</span></span>

|<span data-ttu-id="0f0f1-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="0f0f1-161">Requirement</span></span>| <span data-ttu-id="0f0f1-162">Valor</span><span class="sxs-lookup"><span data-stu-id="0f0f1-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f0f1-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0f0f1-163">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f0f1-164">1.0</span><span class="sxs-lookup"><span data-stu-id="0f0f1-164">1.0</span></span>|
|[<span data-ttu-id="0f0f1-165">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0f0f1-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f0f1-166">Restrito</span><span class="sxs-lookup"><span data-stu-id="0f0f1-166">Restricted</span></span>|
|[<span data-ttu-id="0f0f1-167">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0f0f1-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f0f1-168">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="0f0f1-168">Compose or read</span></span>|