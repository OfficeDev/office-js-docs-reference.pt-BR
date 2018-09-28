# <a name="userprofile"></a><span data-ttu-id="18adc-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="18adc-101">userProfile</span></span>

### <span data-ttu-id="18adc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="18adc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="18adc-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18adc-104">Requirements</span></span>

|<span data-ttu-id="18adc-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="18adc-105">Requirement</span></span>| <span data-ttu-id="18adc-106">Valor</span><span class="sxs-lookup"><span data-stu-id="18adc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="18adc-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18adc-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18adc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="18adc-108">1.0</span></span>|
|[<span data-ttu-id="18adc-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18adc-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18adc-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18adc-110">ReadItem</span></span>|
|[<span data-ttu-id="18adc-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18adc-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18adc-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="18adc-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="18adc-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="18adc-113">Members and methods</span></span>

| <span data-ttu-id="18adc-114">Membro</span><span class="sxs-lookup"><span data-stu-id="18adc-114">Member</span></span> | <span data-ttu-id="18adc-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="18adc-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="18adc-116">displayName</span><span class="sxs-lookup"><span data-stu-id="18adc-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="18adc-117">Membro</span><span class="sxs-lookup"><span data-stu-id="18adc-117">Member</span></span> |
| [<span data-ttu-id="18adc-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="18adc-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="18adc-119">Membro</span><span class="sxs-lookup"><span data-stu-id="18adc-119">Member</span></span> |
| [<span data-ttu-id="18adc-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="18adc-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="18adc-121">Membro</span><span class="sxs-lookup"><span data-stu-id="18adc-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="18adc-122">Membros</span><span class="sxs-lookup"><span data-stu-id="18adc-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="18adc-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="18adc-123">displayName :String</span></span>

<span data-ttu-id="18adc-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="18adc-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="18adc-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18adc-125">Type:</span></span>

*   <span data-ttu-id="18adc-126">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="18adc-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18adc-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18adc-127">Requirements</span></span>

|<span data-ttu-id="18adc-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="18adc-128">Requirement</span></span>| <span data-ttu-id="18adc-129">Valor</span><span class="sxs-lookup"><span data-stu-id="18adc-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="18adc-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18adc-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18adc-131">1.0</span><span class="sxs-lookup"><span data-stu-id="18adc-131">1.0</span></span>|
|[<span data-ttu-id="18adc-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18adc-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18adc-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18adc-133">ReadItem</span></span>|
|[<span data-ttu-id="18adc-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18adc-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18adc-135">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="18adc-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18adc-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18adc-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="18adc-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="18adc-137">emailAddress :String</span></span>

<span data-ttu-id="18adc-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="18adc-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="18adc-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18adc-139">Type:</span></span>

*   <span data-ttu-id="18adc-140">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="18adc-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18adc-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18adc-141">Requirements</span></span>

|<span data-ttu-id="18adc-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="18adc-142">Requirement</span></span>| <span data-ttu-id="18adc-143">Valor</span><span class="sxs-lookup"><span data-stu-id="18adc-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="18adc-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18adc-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18adc-145">1.0</span><span class="sxs-lookup"><span data-stu-id="18adc-145">1.0</span></span>|
|[<span data-ttu-id="18adc-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18adc-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18adc-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18adc-147">ReadItem</span></span>|
|[<span data-ttu-id="18adc-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18adc-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18adc-149">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="18adc-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18adc-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18adc-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="18adc-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="18adc-151">timeZone :String</span></span>

<span data-ttu-id="18adc-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="18adc-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="18adc-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18adc-153">Type:</span></span>

*   <span data-ttu-id="18adc-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="18adc-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18adc-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18adc-155">Requirements</span></span>

|<span data-ttu-id="18adc-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="18adc-156">Requirement</span></span>| <span data-ttu-id="18adc-157">Valor</span><span class="sxs-lookup"><span data-stu-id="18adc-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="18adc-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18adc-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18adc-159">1.0</span><span class="sxs-lookup"><span data-stu-id="18adc-159">1.0</span></span>|
|[<span data-ttu-id="18adc-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18adc-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18adc-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18adc-161">ReadItem</span></span>|
|[<span data-ttu-id="18adc-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18adc-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18adc-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="18adc-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18adc-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18adc-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```