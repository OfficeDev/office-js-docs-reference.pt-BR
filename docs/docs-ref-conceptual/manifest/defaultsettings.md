# <a name="defaultsettings-element"></a><span data-ttu-id="a40f7-101">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="a40f7-101">DefaultSettings element</span></span>

<span data-ttu-id="a40f7-102">Especifica o local de fonte padrão e outras configurações padrão para o seu conteúdo ou adicionar no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a40f7-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="a40f7-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="a40f7-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a40f7-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a40f7-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="a40f7-105">Contidos em</span><span class="sxs-lookup"><span data-stu-id="a40f7-105">Contained in</span></span>

[<span data-ttu-id="a40f7-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a40f7-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="a40f7-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a40f7-107">Can contain</span></span>

|<span data-ttu-id="a40f7-108">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="a40f7-108">**Element**</span></span>|<span data-ttu-id="a40f7-109">**Conteúdo**</span><span class="sxs-lookup"><span data-stu-id="a40f7-109">**Content**</span></span>|<span data-ttu-id="a40f7-110">**Email**</span><span class="sxs-lookup"><span data-stu-id="a40f7-110">**Mail**</span></span>|<span data-ttu-id="a40f7-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a40f7-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a40f7-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a40f7-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="a40f7-113">x</span><span class="sxs-lookup"><span data-stu-id="a40f7-113">x</span></span>||<span data-ttu-id="a40f7-114">x</span><span class="sxs-lookup"><span data-stu-id="a40f7-114">x</span></span>|
|[<span data-ttu-id="a40f7-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="a40f7-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="a40f7-116">x</span><span class="sxs-lookup"><span data-stu-id="a40f7-116">x</span></span>|||
|[<span data-ttu-id="a40f7-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="a40f7-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="a40f7-118">x</span><span class="sxs-lookup"><span data-stu-id="a40f7-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="a40f7-119">Comentários</span><span class="sxs-lookup"><span data-stu-id="a40f7-119">Remarks</span></span>

<span data-ttu-id="a40f7-120">O local de origem e outras configurações no elemento **DefaultSettings** se aplicam apenas a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para os arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="a40f7-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

