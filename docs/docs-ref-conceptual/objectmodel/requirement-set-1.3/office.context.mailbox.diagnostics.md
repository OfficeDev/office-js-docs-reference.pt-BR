
# <a name="diagnostics"></a><span data-ttu-id="80c88-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="80c88-101">diagnostics</span></span>

### <span data-ttu-id="80c88-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="80c88-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="80c88-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="80c88-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="80c88-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="80c88-105">Requirements</span></span>

|<span data-ttu-id="80c88-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="80c88-106">Requirement</span></span>| <span data-ttu-id="80c88-107">Valor</span><span class="sxs-lookup"><span data-stu-id="80c88-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="80c88-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="80c88-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80c88-109">1.0</span><span class="sxs-lookup"><span data-stu-id="80c88-109">1.0</span></span>|
|[<span data-ttu-id="80c88-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="80c88-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80c88-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80c88-111">ReadItem</span></span>|
|[<span data-ttu-id="80c88-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="80c88-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80c88-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="80c88-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="80c88-114">Membros</span><span class="sxs-lookup"><span data-stu-id="80c88-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="80c88-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="80c88-115">hostName :String</span></span>

<span data-ttu-id="80c88-116">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="80c88-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="80c88-117">Uma cadeia de caracteres que pode ser um dos seguintes valores: `Outlook`, `OutlookIOS`, ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="80c88-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="80c88-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="80c88-118">Type:</span></span>

*   <span data-ttu-id="80c88-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="80c88-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80c88-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="80c88-120">Requirements</span></span>

|<span data-ttu-id="80c88-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="80c88-121">Requirement</span></span>| <span data-ttu-id="80c88-122">Valor</span><span class="sxs-lookup"><span data-stu-id="80c88-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="80c88-123">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="80c88-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80c88-124">1.0</span><span class="sxs-lookup"><span data-stu-id="80c88-124">1.0</span></span>|
|[<span data-ttu-id="80c88-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="80c88-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80c88-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80c88-126">ReadItem</span></span>|
|[<span data-ttu-id="80c88-127">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="80c88-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80c88-128">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="80c88-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="80c88-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="80c88-129">hostVersion :String</span></span>

<span data-ttu-id="80c88-130">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="80c88-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="80c88-p102">Se o suplemento de email estiver em execução no cliente do Outlook para área de trabalho ou Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="80c88-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="80c88-134">Tipo:</span><span class="sxs-lookup"><span data-stu-id="80c88-134">Type:</span></span>

*   <span data-ttu-id="80c88-135">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="80c88-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80c88-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="80c88-136">Requirements</span></span>

|<span data-ttu-id="80c88-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="80c88-137">Requirement</span></span>| <span data-ttu-id="80c88-138">Valor</span><span class="sxs-lookup"><span data-stu-id="80c88-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="80c88-139">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="80c88-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80c88-140">1.0</span><span class="sxs-lookup"><span data-stu-id="80c88-140">1.0</span></span>|
|[<span data-ttu-id="80c88-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="80c88-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80c88-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80c88-142">ReadItem</span></span>|
|[<span data-ttu-id="80c88-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="80c88-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80c88-144">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="80c88-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="80c88-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="80c88-145">OWAView :String</span></span>

<span data-ttu-id="80c88-146">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="80c88-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="80c88-147">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="80c88-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="80c88-148">Se o aplicativo host não for Outlook Web App, acessar essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="80c88-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="80c88-149">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e à quantidade de colunas que pode ser exibida:</span><span class="sxs-lookup"><span data-stu-id="80c88-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="80c88-p103">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="80c88-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="80c88-p104">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="80c88-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="80c88-p105">`ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook Web App usa esse modo de exibição em um modo de tela cheia em um computador de mesa.</span><span class="sxs-lookup"><span data-stu-id="80c88-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="80c88-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="80c88-156">Type:</span></span>

*   <span data-ttu-id="80c88-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="80c88-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80c88-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="80c88-158">Requirements</span></span>

|<span data-ttu-id="80c88-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="80c88-159">Requirement</span></span>| <span data-ttu-id="80c88-160">Valor</span><span class="sxs-lookup"><span data-stu-id="80c88-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="80c88-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="80c88-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80c88-162">1.0</span><span class="sxs-lookup"><span data-stu-id="80c88-162">1.0</span></span>|
|[<span data-ttu-id="80c88-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="80c88-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80c88-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80c88-164">ReadItem</span></span>|
|[<span data-ttu-id="80c88-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="80c88-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80c88-166">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="80c88-166">Compose or read</span></span>|