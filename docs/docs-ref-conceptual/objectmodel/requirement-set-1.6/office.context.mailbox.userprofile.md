
# <a name="userprofile"></a><span data-ttu-id="e6e31-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="e6e31-101">userProfile</span></span>

### <span data-ttu-id="e6e31-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="e6e31-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6e31-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e6e31-104">Requirements</span></span>

|<span data-ttu-id="e6e31-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="e6e31-105">Requirement</span></span>| <span data-ttu-id="e6e31-106">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6e31-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e6e31-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6e31-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e6e31-108">1.0</span></span>|
|[<span data-ttu-id="e6e31-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e6e31-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6e31-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6e31-110">ReadItem</span></span>|
|[<span data-ttu-id="e6e31-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e6e31-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6e31-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e6e31-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e6e31-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e6e31-113">Members and methods</span></span>

| <span data-ttu-id="e6e31-114">Membro</span><span class="sxs-lookup"><span data-stu-id="e6e31-114">Member</span></span> | <span data-ttu-id="e6e31-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="e6e31-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e6e31-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e6e31-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e6e31-117">Member</span><span class="sxs-lookup"><span data-stu-id="e6e31-117">Member</span></span> |
| [<span data-ttu-id="e6e31-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e6e31-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e6e31-119">Membro</span><span class="sxs-lookup"><span data-stu-id="e6e31-119">Member</span></span> |
| [<span data-ttu-id="e6e31-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e6e31-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e6e31-121">Membro</span><span class="sxs-lookup"><span data-stu-id="e6e31-121">Member</span></span> |
| [<span data-ttu-id="e6e31-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e6e31-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e6e31-123">Membro</span><span class="sxs-lookup"><span data-stu-id="e6e31-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e6e31-124">Members</span><span class="sxs-lookup"><span data-stu-id="e6e31-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e6e31-125">accountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e6e31-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e6e31-126">Este membro está atualmente apenas com suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="e6e31-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e6e31-127">Obtém o tipo de conta do usuário associado com a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="e6e31-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e6e31-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="e6e31-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e6e31-129">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-129">Value</span></span> | <span data-ttu-id="e6e31-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="e6e31-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e6e31-131">A caixa de correio está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="e6e31-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e6e31-132">A caixa de correio é associada uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="e6e31-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e6e31-133">A caixa de correio está associada com um Office 365 funciona ou escola conta.</span><span class="sxs-lookup"><span data-stu-id="e6e31-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e6e31-134">A caixa de correio é associada uma conta Outlook.com pessoal.</span><span class="sxs-lookup"><span data-stu-id="e6e31-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e6e31-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e6e31-135">Type:</span></span>

*   <span data-ttu-id="e6e31-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e6e31-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6e31-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e6e31-137">Requirements</span></span>

|<span data-ttu-id="e6e31-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="e6e31-138">Requirement</span></span>| <span data-ttu-id="e6e31-139">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6e31-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e6e31-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6e31-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e6e31-141">1.6</span></span> |
|[<span data-ttu-id="e6e31-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e6e31-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6e31-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6e31-143">ReadItem</span></span>|
|[<span data-ttu-id="e6e31-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e6e31-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6e31-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e6e31-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6e31-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e6e31-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="e6e31-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e6e31-147">displayName :String</span></span>

<span data-ttu-id="e6e31-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="e6e31-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e6e31-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e6e31-149">Type:</span></span>

*   <span data-ttu-id="e6e31-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e6e31-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6e31-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e6e31-151">Requirements</span></span>

|<span data-ttu-id="e6e31-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="e6e31-152">Requirement</span></span>| <span data-ttu-id="e6e31-153">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6e31-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e6e31-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6e31-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e6e31-155">1.0</span></span>|
|[<span data-ttu-id="e6e31-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e6e31-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6e31-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6e31-157">ReadItem</span></span>|
|[<span data-ttu-id="e6e31-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e6e31-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6e31-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e6e31-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6e31-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e6e31-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e6e31-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e6e31-161">emailAddress :String</span></span>

<span data-ttu-id="e6e31-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="e6e31-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e6e31-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e6e31-163">Type:</span></span>

*   <span data-ttu-id="e6e31-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e6e31-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6e31-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e6e31-165">Requirements</span></span>

|<span data-ttu-id="e6e31-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="e6e31-166">Requirement</span></span>| <span data-ttu-id="e6e31-167">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6e31-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e6e31-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6e31-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e6e31-169">1.0</span></span>|
|[<span data-ttu-id="e6e31-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e6e31-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6e31-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6e31-171">ReadItem</span></span>|
|[<span data-ttu-id="e6e31-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e6e31-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6e31-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e6e31-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6e31-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e6e31-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e6e31-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e6e31-175">timeZone :String</span></span>

<span data-ttu-id="e6e31-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="e6e31-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e6e31-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e6e31-177">Type:</span></span>

*   <span data-ttu-id="e6e31-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e6e31-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6e31-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e6e31-179">Requirements</span></span>

|<span data-ttu-id="e6e31-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="e6e31-180">Requirement</span></span>| <span data-ttu-id="e6e31-181">Valor</span><span class="sxs-lookup"><span data-stu-id="e6e31-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6e31-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e6e31-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6e31-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e6e31-183">1.0</span></span>|
|[<span data-ttu-id="e6e31-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e6e31-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6e31-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6e31-185">ReadItem</span></span>|
|[<span data-ttu-id="e6e31-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e6e31-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6e31-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="e6e31-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6e31-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e6e31-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```