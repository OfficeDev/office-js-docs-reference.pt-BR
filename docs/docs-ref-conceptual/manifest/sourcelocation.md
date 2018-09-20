# <a name="sourcelocation-element"></a><span data-ttu-id="68f63-101">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="68f63-101">SourceLocation element</span></span>

<span data-ttu-id="68f63-p101">Especifica o local de origem do arquivo para o Suplemento do Office como uma URL que contém entre 1 e 2.018 caracteres de comprimento. O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.</span><span class="sxs-lookup"><span data-stu-id="68f63-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="68f63-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="68f63-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="68f63-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="68f63-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="68f63-106">Contidos em</span><span class="sxs-lookup"><span data-stu-id="68f63-106">Contained in</span></span>

- <span data-ttu-id="68f63-107">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="68f63-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="68f63-108">[FormSettings](formsettings.md) (suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="68f63-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="68f63-109">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email)</span><span class="sxs-lookup"><span data-stu-id="68f63-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="68f63-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="68f63-110">Can contain</span></span>

[<span data-ttu-id="68f63-111">Override</span><span class="sxs-lookup"><span data-stu-id="68f63-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="68f63-112">Atributos</span><span class="sxs-lookup"><span data-stu-id="68f63-112">Attributes</span></span>

|<span data-ttu-id="68f63-113">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="68f63-113">**Attribute**</span></span>|<span data-ttu-id="68f63-114">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="68f63-114">**Type**</span></span>|<span data-ttu-id="68f63-115">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="68f63-115">**Required**</span></span>|<span data-ttu-id="68f63-116">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="68f63-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="68f63-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="68f63-117">DefaultValue</span></span>|<span data-ttu-id="68f63-118">URL</span><span class="sxs-lookup"><span data-stu-id="68f63-118">URL</span></span>|<span data-ttu-id="68f63-119">obrigatório</span><span class="sxs-lookup"><span data-stu-id="68f63-119">required</span></span>|<span data-ttu-id="68f63-120">Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="68f63-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
