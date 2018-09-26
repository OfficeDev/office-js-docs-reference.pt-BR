
# <a name="userprofile"></a><span data-ttu-id="fde84-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="fde84-101">userProfile</span></span>

### <span data-ttu-id="fde84-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="fde84-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="fde84-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fde84-104">Requirements</span></span>

|<span data-ttu-id="fde84-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="fde84-105">Requirement</span></span>| <span data-ttu-id="fde84-106">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="fde84-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fde84-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fde84-108">1.0</span><span class="sxs-lookup"><span data-stu-id="fde84-108">1.0</span></span>|
|[<span data-ttu-id="fde84-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fde84-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fde84-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fde84-110">ReadItem</span></span>|
|[<span data-ttu-id="fde84-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fde84-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fde84-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fde84-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fde84-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="fde84-113">Members and methods</span></span>

| <span data-ttu-id="fde84-114">Membro</span><span class="sxs-lookup"><span data-stu-id="fde84-114">Member</span></span> | <span data-ttu-id="fde84-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="fde84-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fde84-116">accountType</span><span class="sxs-lookup"><span data-stu-id="fde84-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="fde84-117">Member</span><span class="sxs-lookup"><span data-stu-id="fde84-117">Member</span></span> |
| [<span data-ttu-id="fde84-118">displayName</span><span class="sxs-lookup"><span data-stu-id="fde84-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="fde84-119">Membro</span><span class="sxs-lookup"><span data-stu-id="fde84-119">Member</span></span> |
| [<span data-ttu-id="fde84-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="fde84-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="fde84-121">Membro</span><span class="sxs-lookup"><span data-stu-id="fde84-121">Member</span></span> |
| [<span data-ttu-id="fde84-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="fde84-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="fde84-123">Membro</span><span class="sxs-lookup"><span data-stu-id="fde84-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="fde84-124">Members</span><span class="sxs-lookup"><span data-stu-id="fde84-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="fde84-125">accountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fde84-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="fde84-126">Este membro está atualmente apenas com suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="fde84-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="fde84-127">Obtém o tipo de conta do usuário associado com a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="fde84-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="fde84-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="fde84-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="fde84-129">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-129">Value</span></span> | <span data-ttu-id="fde84-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="fde84-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="fde84-131">A caixa de correio está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="fde84-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="fde84-132">A caixa de correio é associada uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="fde84-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="fde84-133">A caixa de correio está associada com um Office 365 funciona ou escola conta.</span><span class="sxs-lookup"><span data-stu-id="fde84-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="fde84-134">A caixa de correio é associada uma conta Outlook.com pessoal.</span><span class="sxs-lookup"><span data-stu-id="fde84-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="fde84-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fde84-135">Type:</span></span>

*   <span data-ttu-id="fde84-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fde84-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fde84-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fde84-137">Requirements</span></span>

|<span data-ttu-id="fde84-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="fde84-138">Requirement</span></span>| <span data-ttu-id="fde84-139">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="fde84-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fde84-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fde84-141">1.6</span><span class="sxs-lookup"><span data-stu-id="fde84-141">1.6</span></span> |
|[<span data-ttu-id="fde84-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fde84-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fde84-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fde84-143">ReadItem</span></span>|
|[<span data-ttu-id="fde84-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fde84-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fde84-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fde84-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fde84-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fde84-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="fde84-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="fde84-147">displayName :String</span></span>

<span data-ttu-id="fde84-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="fde84-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="fde84-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fde84-149">Type:</span></span>

*   <span data-ttu-id="fde84-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fde84-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fde84-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fde84-151">Requirements</span></span>

|<span data-ttu-id="fde84-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="fde84-152">Requirement</span></span>| <span data-ttu-id="fde84-153">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="fde84-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fde84-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fde84-155">1.0</span><span class="sxs-lookup"><span data-stu-id="fde84-155">1.0</span></span>|
|[<span data-ttu-id="fde84-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fde84-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fde84-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fde84-157">ReadItem</span></span>|
|[<span data-ttu-id="fde84-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fde84-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fde84-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fde84-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fde84-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fde84-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="fde84-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="fde84-161">emailAddress :String</span></span>

<span data-ttu-id="fde84-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="fde84-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="fde84-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fde84-163">Type:</span></span>

*   <span data-ttu-id="fde84-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fde84-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fde84-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fde84-165">Requirements</span></span>

|<span data-ttu-id="fde84-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="fde84-166">Requirement</span></span>| <span data-ttu-id="fde84-167">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="fde84-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fde84-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fde84-169">1.0</span><span class="sxs-lookup"><span data-stu-id="fde84-169">1.0</span></span>|
|[<span data-ttu-id="fde84-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fde84-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fde84-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fde84-171">ReadItem</span></span>|
|[<span data-ttu-id="fde84-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fde84-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fde84-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fde84-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fde84-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fde84-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="fde84-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="fde84-175">timeZone :String</span></span>

<span data-ttu-id="fde84-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="fde84-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="fde84-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fde84-177">Type:</span></span>

*   <span data-ttu-id="fde84-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fde84-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fde84-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fde84-179">Requirements</span></span>

|<span data-ttu-id="fde84-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="fde84-180">Requirement</span></span>| <span data-ttu-id="fde84-181">Valor</span><span class="sxs-lookup"><span data-stu-id="fde84-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="fde84-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fde84-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fde84-183">1.0</span><span class="sxs-lookup"><span data-stu-id="fde84-183">1.0</span></span>|
|[<span data-ttu-id="fde84-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fde84-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fde84-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fde84-185">ReadItem</span></span>|
|[<span data-ttu-id="fde84-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fde84-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fde84-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fde84-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fde84-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fde84-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```