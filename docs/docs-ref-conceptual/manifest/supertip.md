# <a name="supertip"></a><span data-ttu-id="ae657-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="ae657-101">Supertip</span></span>

<span data-ttu-id="ae657-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="ae657-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ae657-104">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ae657-104">Child elements</span></span>

|  <span data-ttu-id="ae657-105">Elemento</span><span class="sxs-lookup"><span data-stu-id="ae657-105">Element</span></span> |  <span data-ttu-id="ae657-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ae657-106">Required</span></span>  |  <span data-ttu-id="ae657-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae657-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ae657-108">Título</span><span class="sxs-lookup"><span data-stu-id="ae657-108">Title</span></span>](#title)        | <span data-ttu-id="ae657-109">Sim</span><span class="sxs-lookup"><span data-stu-id="ae657-109">Yes</span></span> |   <span data-ttu-id="ae657-110">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="ae657-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="ae657-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae657-111">Description</span></span>](#description)  | <span data-ttu-id="ae657-112">Sim</span><span class="sxs-lookup"><span data-stu-id="ae657-112">Yes</span></span> |  <span data-ttu-id="ae657-113">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="ae657-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="ae657-114">Title</span><span class="sxs-lookup"><span data-stu-id="ae657-114">Title</span></span>

<span data-ttu-id="ae657-p102">Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ae657-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="ae657-118">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae657-118">Description</span></span>

<span data-ttu-id="ae657-p103">Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **LongStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ae657-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="ae657-122">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ae657-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
