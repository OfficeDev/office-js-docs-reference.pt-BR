
# <a name="userprofile"></a><span data-ttu-id="236ec-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="236ec-101">userProfile</span></span>

### <span data-ttu-id="236ec-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="236ec-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="236ec-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="236ec-104">Requirements</span></span>

|<span data-ttu-id="236ec-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="236ec-105">Requirement</span></span>| <span data-ttu-id="236ec-106">Valor</span><span class="sxs-lookup"><span data-stu-id="236ec-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="236ec-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="236ec-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236ec-108">1.0</span><span class="sxs-lookup"><span data-stu-id="236ec-108">1.0</span></span>|
|[<span data-ttu-id="236ec-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="236ec-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236ec-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236ec-110">ReadItem</span></span>|
|[<span data-ttu-id="236ec-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="236ec-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="236ec-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="236ec-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="236ec-113">Membros</span><span class="sxs-lookup"><span data-stu-id="236ec-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="236ec-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="236ec-114">displayName :String</span></span>

<span data-ttu-id="236ec-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="236ec-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="236ec-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="236ec-116">Type:</span></span>

*   <span data-ttu-id="236ec-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="236ec-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="236ec-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="236ec-118">Requirements</span></span>

|<span data-ttu-id="236ec-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="236ec-119">Requirement</span></span>| <span data-ttu-id="236ec-120">Valor</span><span class="sxs-lookup"><span data-stu-id="236ec-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="236ec-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="236ec-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236ec-122">1.0</span><span class="sxs-lookup"><span data-stu-id="236ec-122">1.0</span></span>|
|[<span data-ttu-id="236ec-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="236ec-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236ec-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236ec-124">ReadItem</span></span>|
|[<span data-ttu-id="236ec-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="236ec-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="236ec-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="236ec-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="236ec-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="236ec-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="236ec-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="236ec-128">emailAddress :String</span></span>

<span data-ttu-id="236ec-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="236ec-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="236ec-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="236ec-130">Type:</span></span>

*   <span data-ttu-id="236ec-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="236ec-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="236ec-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="236ec-132">Requirements</span></span>

|<span data-ttu-id="236ec-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="236ec-133">Requirement</span></span>| <span data-ttu-id="236ec-134">Valor</span><span class="sxs-lookup"><span data-stu-id="236ec-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="236ec-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="236ec-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236ec-136">1.0</span><span class="sxs-lookup"><span data-stu-id="236ec-136">1.0</span></span>|
|[<span data-ttu-id="236ec-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="236ec-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236ec-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236ec-138">ReadItem</span></span>|
|[<span data-ttu-id="236ec-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="236ec-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="236ec-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="236ec-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="236ec-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="236ec-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="236ec-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="236ec-142">timeZone :String</span></span>

<span data-ttu-id="236ec-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="236ec-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="236ec-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="236ec-144">Type:</span></span>

*   <span data-ttu-id="236ec-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="236ec-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="236ec-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="236ec-146">Requirements</span></span>

|<span data-ttu-id="236ec-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="236ec-147">Requirement</span></span>| <span data-ttu-id="236ec-148">Valor</span><span class="sxs-lookup"><span data-stu-id="236ec-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="236ec-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="236ec-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236ec-150">1.0</span><span class="sxs-lookup"><span data-stu-id="236ec-150">1.0</span></span>|
|[<span data-ttu-id="236ec-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="236ec-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236ec-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236ec-152">ReadItem</span></span>|
|[<span data-ttu-id="236ec-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="236ec-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="236ec-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="236ec-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="236ec-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="236ec-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```