
# <a name="userprofile"></a><span data-ttu-id="fbf8f-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="fbf8f-101">userProfile</span></span>

### <span data-ttu-id="fbf8f-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="fbf8f-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbf8f-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-104">Requirements</span></span>

|<span data-ttu-id="fbf8f-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="fbf8f-105">Requirement</span></span>| <span data-ttu-id="fbf8f-106">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbf8f-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fbf8f-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbf8f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="fbf8f-108">1.0</span></span>|
|[<span data-ttu-id="fbf8f-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbf8f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbf8f-110">ReadItem</span></span>|
|[<span data-ttu-id="fbf8f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fbf8f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbf8f-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fbf8f-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fbf8f-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-113">Members and methods</span></span>

| <span data-ttu-id="fbf8f-114">Membro</span><span class="sxs-lookup"><span data-stu-id="fbf8f-114">Member</span></span> | <span data-ttu-id="fbf8f-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fbf8f-116">accountType</span><span class="sxs-lookup"><span data-stu-id="fbf8f-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="fbf8f-117">Member</span><span class="sxs-lookup"><span data-stu-id="fbf8f-117">Member</span></span> |
| [<span data-ttu-id="fbf8f-118">displayName</span><span class="sxs-lookup"><span data-stu-id="fbf8f-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="fbf8f-119">Membro</span><span class="sxs-lookup"><span data-stu-id="fbf8f-119">Member</span></span> |
| [<span data-ttu-id="fbf8f-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="fbf8f-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="fbf8f-121">Membro</span><span class="sxs-lookup"><span data-stu-id="fbf8f-121">Member</span></span> |
| [<span data-ttu-id="fbf8f-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="fbf8f-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="fbf8f-123">Membro</span><span class="sxs-lookup"><span data-stu-id="fbf8f-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="fbf8f-124">Members</span><span class="sxs-lookup"><span data-stu-id="fbf8f-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="fbf8f-125">accountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fbf8f-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="fbf8f-126">Este membro está atualmente apenas 2016 com suporte no Outlook para Mac, 16.9.1212 de compilação e maiores.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="fbf8f-127">Obtém o tipo de conta do usuário associado com a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="fbf8f-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="fbf8f-129">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-129">Value</span></span> | <span data-ttu-id="fbf8f-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="fbf8f-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="fbf8f-131">A caixa de correio está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="fbf8f-132">A caixa de correio é associada uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="fbf8f-133">A caixa de correio está associada com um Office 365 funciona ou escola conta.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="fbf8f-134">A caixa de correio é associada uma conta Outlook.com pessoal.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="fbf8f-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fbf8f-135">Type:</span></span>

*   <span data-ttu-id="fbf8f-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fbf8f-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbf8f-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-137">Requirements</span></span>

|<span data-ttu-id="fbf8f-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="fbf8f-138">Requirement</span></span>| <span data-ttu-id="fbf8f-139">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbf8f-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fbf8f-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbf8f-141">1.6</span><span class="sxs-lookup"><span data-stu-id="fbf8f-141">1.6</span></span> |
|[<span data-ttu-id="fbf8f-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbf8f-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbf8f-143">ReadItem</span></span>|
|[<span data-ttu-id="fbf8f-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fbf8f-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbf8f-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fbf8f-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbf8f-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="fbf8f-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="fbf8f-147">displayName :String</span></span>

<span data-ttu-id="fbf8f-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="fbf8f-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fbf8f-149">Type:</span></span>

*   <span data-ttu-id="fbf8f-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fbf8f-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbf8f-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-151">Requirements</span></span>

|<span data-ttu-id="fbf8f-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="fbf8f-152">Requirement</span></span>| <span data-ttu-id="fbf8f-153">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbf8f-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fbf8f-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbf8f-155">1.0</span><span class="sxs-lookup"><span data-stu-id="fbf8f-155">1.0</span></span>|
|[<span data-ttu-id="fbf8f-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbf8f-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbf8f-157">ReadItem</span></span>|
|[<span data-ttu-id="fbf8f-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fbf8f-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbf8f-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fbf8f-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbf8f-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="fbf8f-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="fbf8f-161">emailAddress :String</span></span>

<span data-ttu-id="fbf8f-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="fbf8f-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fbf8f-163">Type:</span></span>

*   <span data-ttu-id="fbf8f-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fbf8f-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbf8f-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-165">Requirements</span></span>

|<span data-ttu-id="fbf8f-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="fbf8f-166">Requirement</span></span>| <span data-ttu-id="fbf8f-167">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbf8f-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fbf8f-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbf8f-169">1.0</span><span class="sxs-lookup"><span data-stu-id="fbf8f-169">1.0</span></span>|
|[<span data-ttu-id="fbf8f-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbf8f-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbf8f-171">ReadItem</span></span>|
|[<span data-ttu-id="fbf8f-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fbf8f-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbf8f-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fbf8f-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbf8f-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="fbf8f-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="fbf8f-175">timeZone :String</span></span>

<span data-ttu-id="fbf8f-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="fbf8f-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="fbf8f-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fbf8f-177">Type:</span></span>

*   <span data-ttu-id="fbf8f-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fbf8f-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbf8f-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fbf8f-179">Requirements</span></span>

|<span data-ttu-id="fbf8f-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="fbf8f-180">Requirement</span></span>| <span data-ttu-id="fbf8f-181">Valor</span><span class="sxs-lookup"><span data-stu-id="fbf8f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbf8f-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fbf8f-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbf8f-183">1.0</span><span class="sxs-lookup"><span data-stu-id="fbf8f-183">1.0</span></span>|
|[<span data-ttu-id="fbf8f-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbf8f-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbf8f-185">ReadItem</span></span>|
|[<span data-ttu-id="fbf8f-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fbf8f-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbf8f-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fbf8f-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbf8f-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fbf8f-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```