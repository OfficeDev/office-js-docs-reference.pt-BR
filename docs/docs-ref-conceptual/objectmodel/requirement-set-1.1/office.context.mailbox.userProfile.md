
# <a name="userprofile"></a><span data-ttu-id="b2cf6-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b2cf6-101">userProfile</span></span>

### <span data-ttu-id="b2cf6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b2cf6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2cf6-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b2cf6-104">Requirements</span></span>

|<span data-ttu-id="b2cf6-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b2cf6-105">Requirement</span></span>| <span data-ttu-id="b2cf6-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b2cf6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2cf6-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b2cf6-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2cf6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b2cf6-108">1.0</span></span>|
|[<span data-ttu-id="b2cf6-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2cf6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2cf6-110">ReadItem</span></span>|
|[<span data-ttu-id="b2cf6-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b2cf6-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2cf6-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b2cf6-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="b2cf6-113">Membros</span><span class="sxs-lookup"><span data-stu-id="b2cf6-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b2cf6-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b2cf6-114">displayName :String</span></span>

<span data-ttu-id="b2cf6-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="b2cf6-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b2cf6-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b2cf6-116">Type:</span></span>

*   <span data-ttu-id="b2cf6-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b2cf6-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2cf6-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b2cf6-118">Requirements</span></span>

|<span data-ttu-id="b2cf6-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="b2cf6-119">Requirement</span></span>| <span data-ttu-id="b2cf6-120">Valor</span><span class="sxs-lookup"><span data-stu-id="b2cf6-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2cf6-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b2cf6-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2cf6-122">1.0</span><span class="sxs-lookup"><span data-stu-id="b2cf6-122">1.0</span></span>|
|[<span data-ttu-id="b2cf6-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2cf6-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2cf6-124">ReadItem</span></span>|
|[<span data-ttu-id="b2cf6-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b2cf6-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2cf6-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b2cf6-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2cf6-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b2cf6-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b2cf6-128">emailAddress :String</span></span>

<span data-ttu-id="b2cf6-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="b2cf6-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b2cf6-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b2cf6-130">Type:</span></span>

*   <span data-ttu-id="b2cf6-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b2cf6-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2cf6-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b2cf6-132">Requirements</span></span>

|<span data-ttu-id="b2cf6-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="b2cf6-133">Requirement</span></span>| <span data-ttu-id="b2cf6-134">Valor</span><span class="sxs-lookup"><span data-stu-id="b2cf6-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2cf6-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b2cf6-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2cf6-136">1.0</span><span class="sxs-lookup"><span data-stu-id="b2cf6-136">1.0</span></span>|
|[<span data-ttu-id="b2cf6-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2cf6-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2cf6-138">ReadItem</span></span>|
|[<span data-ttu-id="b2cf6-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b2cf6-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2cf6-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b2cf6-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2cf6-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b2cf6-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b2cf6-142">timeZone :String</span></span>

<span data-ttu-id="b2cf6-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="b2cf6-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b2cf6-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b2cf6-144">Type:</span></span>

*   <span data-ttu-id="b2cf6-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b2cf6-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2cf6-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b2cf6-146">Requirements</span></span>

|<span data-ttu-id="b2cf6-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="b2cf6-147">Requirement</span></span>| <span data-ttu-id="b2cf6-148">Valor</span><span class="sxs-lookup"><span data-stu-id="b2cf6-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2cf6-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b2cf6-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2cf6-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b2cf6-150">1.0</span></span>|
|[<span data-ttu-id="b2cf6-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2cf6-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2cf6-152">ReadItem</span></span>|
|[<span data-ttu-id="b2cf6-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b2cf6-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2cf6-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b2cf6-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2cf6-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b2cf6-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```