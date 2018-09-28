
# <a name="userprofile"></a><span data-ttu-id="6a879-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="6a879-101">userProfile</span></span>

### <span data-ttu-id="6a879-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="6a879-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a879-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a879-104">Requirements</span></span>

|<span data-ttu-id="6a879-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a879-105">Requirement</span></span>| <span data-ttu-id="6a879-106">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a879-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a879-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6a879-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6a879-108">1.0</span></span>|
|[<span data-ttu-id="6a879-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6a879-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6a879-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6a879-110">ReadItem</span></span>|
|[<span data-ttu-id="6a879-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a879-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6a879-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="6a879-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6a879-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="6a879-113">Members and methods</span></span>

| <span data-ttu-id="6a879-114">Membro</span><span class="sxs-lookup"><span data-stu-id="6a879-114">Member</span></span> | <span data-ttu-id="6a879-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a879-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6a879-116">accountType</span><span class="sxs-lookup"><span data-stu-id="6a879-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="6a879-117">Member</span><span class="sxs-lookup"><span data-stu-id="6a879-117">Member</span></span> |
| [<span data-ttu-id="6a879-118">displayName</span><span class="sxs-lookup"><span data-stu-id="6a879-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="6a879-119">Membro</span><span class="sxs-lookup"><span data-stu-id="6a879-119">Member</span></span> |
| [<span data-ttu-id="6a879-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6a879-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6a879-121">Membro</span><span class="sxs-lookup"><span data-stu-id="6a879-121">Member</span></span> |
| [<span data-ttu-id="6a879-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="6a879-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6a879-123">Membro</span><span class="sxs-lookup"><span data-stu-id="6a879-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="6a879-124">Members</span><span class="sxs-lookup"><span data-stu-id="6a879-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="6a879-125">accountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a879-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="6a879-126">Este membro está atualmente apenas 2016 com suporte no Outlook para Mac, 16.9.1212 de compilação e maiores.</span><span class="sxs-lookup"><span data-stu-id="6a879-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="6a879-127">Obtém o tipo de conta do usuário associado com a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="6a879-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="6a879-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a879-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="6a879-129">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-129">Value</span></span> | <span data-ttu-id="6a879-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="6a879-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="6a879-131">A caixa de correio está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="6a879-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="6a879-132">A caixa de correio é associada uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="6a879-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="6a879-133">A caixa de correio está associada com um Office 365 funciona ou escola conta.</span><span class="sxs-lookup"><span data-stu-id="6a879-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="6a879-134">A caixa de correio é associada uma conta Outlook.com pessoal.</span><span class="sxs-lookup"><span data-stu-id="6a879-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="6a879-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6a879-135">Type:</span></span>

*   <span data-ttu-id="6a879-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a879-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a879-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a879-137">Requirements</span></span>

|<span data-ttu-id="6a879-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a879-138">Requirement</span></span>| <span data-ttu-id="6a879-139">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a879-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a879-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6a879-141">1.6</span><span class="sxs-lookup"><span data-stu-id="6a879-141">1.6</span></span> |
|[<span data-ttu-id="6a879-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6a879-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6a879-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6a879-143">ReadItem</span></span>|
|[<span data-ttu-id="6a879-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a879-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6a879-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="6a879-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6a879-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6a879-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="6a879-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="6a879-147">displayName :String</span></span>

<span data-ttu-id="6a879-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="6a879-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6a879-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6a879-149">Type:</span></span>

*   <span data-ttu-id="6a879-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a879-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a879-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a879-151">Requirements</span></span>

|<span data-ttu-id="6a879-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a879-152">Requirement</span></span>| <span data-ttu-id="6a879-153">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a879-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a879-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6a879-155">1.0</span><span class="sxs-lookup"><span data-stu-id="6a879-155">1.0</span></span>|
|[<span data-ttu-id="6a879-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6a879-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6a879-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6a879-157">ReadItem</span></span>|
|[<span data-ttu-id="6a879-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a879-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6a879-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="6a879-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6a879-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6a879-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="6a879-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="6a879-161">emailAddress :String</span></span>

<span data-ttu-id="6a879-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="6a879-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6a879-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6a879-163">Type:</span></span>

*   <span data-ttu-id="6a879-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a879-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a879-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a879-165">Requirements</span></span>

|<span data-ttu-id="6a879-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a879-166">Requirement</span></span>| <span data-ttu-id="6a879-167">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a879-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a879-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6a879-169">1.0</span><span class="sxs-lookup"><span data-stu-id="6a879-169">1.0</span></span>|
|[<span data-ttu-id="6a879-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6a879-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6a879-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6a879-171">ReadItem</span></span>|
|[<span data-ttu-id="6a879-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a879-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6a879-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="6a879-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6a879-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6a879-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="6a879-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="6a879-175">timeZone :String</span></span>

<span data-ttu-id="6a879-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="6a879-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6a879-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6a879-177">Type:</span></span>

*   <span data-ttu-id="6a879-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a879-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a879-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a879-179">Requirements</span></span>

|<span data-ttu-id="6a879-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a879-180">Requirement</span></span>| <span data-ttu-id="6a879-181">Valor</span><span class="sxs-lookup"><span data-stu-id="6a879-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a879-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a879-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6a879-183">1.0</span><span class="sxs-lookup"><span data-stu-id="6a879-183">1.0</span></span>|
|[<span data-ttu-id="6a879-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6a879-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6a879-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6a879-185">ReadItem</span></span>|
|[<span data-ttu-id="6a879-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a879-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6a879-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="6a879-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6a879-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6a879-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```