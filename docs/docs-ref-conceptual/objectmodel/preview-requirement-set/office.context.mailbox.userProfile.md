
# <a name="userprofile"></a><span data-ttu-id="8f089-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="8f089-101">userProfile</span></span>

### <span data-ttu-id="8f089-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="8f089-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f089-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f089-104">Requirements</span></span>

|<span data-ttu-id="8f089-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f089-105">Requirement</span></span>| <span data-ttu-id="8f089-106">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f089-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f089-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f089-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8f089-108">1.0</span></span>|
|[<span data-ttu-id="8f089-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f089-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8f089-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8f089-110">ReadItem</span></span>|
|[<span data-ttu-id="8f089-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f089-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f089-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8f089-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8f089-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8f089-113">Members and methods</span></span>

| <span data-ttu-id="8f089-114">Membro</span><span class="sxs-lookup"><span data-stu-id="8f089-114">Member</span></span> | <span data-ttu-id="8f089-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f089-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8f089-116">accountType</span><span class="sxs-lookup"><span data-stu-id="8f089-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="8f089-117">Member</span><span class="sxs-lookup"><span data-stu-id="8f089-117">Member</span></span> |
| [<span data-ttu-id="8f089-118">displayName</span><span class="sxs-lookup"><span data-stu-id="8f089-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="8f089-119">Membro</span><span class="sxs-lookup"><span data-stu-id="8f089-119">Member</span></span> |
| [<span data-ttu-id="8f089-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8f089-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8f089-121">Membro</span><span class="sxs-lookup"><span data-stu-id="8f089-121">Member</span></span> |
| [<span data-ttu-id="8f089-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="8f089-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8f089-123">Membro</span><span class="sxs-lookup"><span data-stu-id="8f089-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8f089-124">Members</span><span class="sxs-lookup"><span data-stu-id="8f089-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="8f089-125">accountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f089-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="8f089-126">Este membro está atualmente apenas com suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="8f089-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="8f089-127">Obtém o tipo de conta do usuário associado com a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8f089-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="8f089-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8f089-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="8f089-129">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-129">Value</span></span> | <span data-ttu-id="8f089-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="8f089-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="8f089-131">A caixa de correio está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="8f089-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="8f089-132">A caixa de correio é associada uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="8f089-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="8f089-133">A caixa de correio está associada com um Office 365 funciona ou escola conta.</span><span class="sxs-lookup"><span data-stu-id="8f089-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="8f089-134">A caixa de correio é associada uma conta Outlook.com pessoal.</span><span class="sxs-lookup"><span data-stu-id="8f089-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="8f089-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8f089-135">Type:</span></span>

*   <span data-ttu-id="8f089-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f089-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f089-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f089-137">Requirements</span></span>

|<span data-ttu-id="8f089-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f089-138">Requirement</span></span>| <span data-ttu-id="8f089-139">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f089-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f089-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f089-141">1.6</span><span class="sxs-lookup"><span data-stu-id="8f089-141">1.6</span></span> |
|[<span data-ttu-id="8f089-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f089-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8f089-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8f089-143">ReadItem</span></span>|
|[<span data-ttu-id="8f089-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f089-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f089-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8f089-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f089-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f089-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="8f089-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="8f089-147">displayName :String</span></span>

<span data-ttu-id="8f089-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="8f089-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8f089-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8f089-149">Type:</span></span>

*   <span data-ttu-id="8f089-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f089-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f089-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f089-151">Requirements</span></span>

|<span data-ttu-id="8f089-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f089-152">Requirement</span></span>| <span data-ttu-id="8f089-153">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f089-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f089-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f089-155">1.0</span><span class="sxs-lookup"><span data-stu-id="8f089-155">1.0</span></span>|
|[<span data-ttu-id="8f089-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f089-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8f089-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8f089-157">ReadItem</span></span>|
|[<span data-ttu-id="8f089-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f089-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f089-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8f089-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f089-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f089-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="8f089-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="8f089-161">emailAddress :String</span></span>

<span data-ttu-id="8f089-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="8f089-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8f089-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8f089-163">Type:</span></span>

*   <span data-ttu-id="8f089-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f089-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f089-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f089-165">Requirements</span></span>

|<span data-ttu-id="8f089-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f089-166">Requirement</span></span>| <span data-ttu-id="8f089-167">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f089-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f089-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f089-169">1.0</span><span class="sxs-lookup"><span data-stu-id="8f089-169">1.0</span></span>|
|[<span data-ttu-id="8f089-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f089-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8f089-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8f089-171">ReadItem</span></span>|
|[<span data-ttu-id="8f089-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f089-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f089-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8f089-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f089-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f089-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="8f089-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="8f089-175">timeZone :String</span></span>

<span data-ttu-id="8f089-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="8f089-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8f089-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8f089-177">Type:</span></span>

*   <span data-ttu-id="8f089-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f089-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f089-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f089-179">Requirements</span></span>

|<span data-ttu-id="8f089-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f089-180">Requirement</span></span>| <span data-ttu-id="8f089-181">Valor</span><span class="sxs-lookup"><span data-stu-id="8f089-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f089-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f089-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f089-183">1.0</span><span class="sxs-lookup"><span data-stu-id="8f089-183">1.0</span></span>|
|[<span data-ttu-id="8f089-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f089-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8f089-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8f089-185">ReadItem</span></span>|
|[<span data-ttu-id="8f089-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f089-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f089-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8f089-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f089-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f089-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```