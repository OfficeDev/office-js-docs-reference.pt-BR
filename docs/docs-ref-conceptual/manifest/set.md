# <a name="set-element"></a><span data-ttu-id="ce487-101">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="ce487-101">Set element</span></span>

<span data-ttu-id="ce487-102">Especifica um conjunto de requisitos a partir da API do JavaScript para Office que o seu Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="ce487-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="ce487-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="ce487-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ce487-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ce487-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="ce487-105">Contidos em</span><span class="sxs-lookup"><span data-stu-id="ce487-105">Contained in</span></span>

[<span data-ttu-id="ce487-106">Sets</span><span class="sxs-lookup"><span data-stu-id="ce487-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="ce487-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="ce487-107">Attributes</span></span>

|<span data-ttu-id="ce487-108">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="ce487-108">**Attribute**</span></span>|<span data-ttu-id="ce487-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="ce487-109">**Type**</span></span>|<span data-ttu-id="ce487-110">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="ce487-110">**Required**</span></span>|<span data-ttu-id="ce487-111">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ce487-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ce487-112">Nome</span><span class="sxs-lookup"><span data-stu-id="ce487-112">Name</span></span>|<span data-ttu-id="ce487-113">string</span><span class="sxs-lookup"><span data-stu-id="ce487-113">string</span></span>|<span data-ttu-id="ce487-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ce487-114">required</span></span>|<span data-ttu-id="ce487-115">O nome de um [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce487-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="ce487-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="ce487-116">MinVersion</span></span>|<span data-ttu-id="ce487-117">string</span><span class="sxs-lookup"><span data-stu-id="ce487-117">string</span></span>|<span data-ttu-id="ce487-118">opcional</span><span class="sxs-lookup"><span data-stu-id="ce487-118">optional</span></span>|<span data-ttu-id="ce487-p101">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="ce487-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="ce487-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="ce487-121">Remarks</span></span>

<span data-ttu-id="ce487-122">Para obter mais informações sobre os conjuntos de requisito, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce487-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="ce487-123">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="ce487-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ce487-124">Para suplementos de email, há somente um `"Mailbox"` requisito conjunto disponível.</span><span class="sxs-lookup"><span data-stu-id="ce487-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="ce487-125">Esse conjunto de requisito contém o subconjunto inteiro de API suportada em suplementos de email para Outlook, e você deve especificar o `"Mailbox"` requisito definir no manifesto do seu email suplemento (não é opcional conforme for o caso do conteúdo e tarefa suplementos do painel).</span><span class="sxs-lookup"><span data-stu-id="ce487-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="ce487-126">Além disso, você não pode declarar suporte para métodos específicos em suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="ce487-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
