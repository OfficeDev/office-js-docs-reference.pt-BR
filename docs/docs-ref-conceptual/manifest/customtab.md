# <a name="customtab-element"></a><span data-ttu-id="5919d-101">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="5919d-101">CustomTab element</span></span>

<span data-ttu-id="5919d-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="5919d-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="5919d-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="5919d-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="5919d-107">O atributo **id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5919d-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5919d-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="5919d-108">Child elements</span></span>

|  <span data-ttu-id="5919d-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="5919d-109">Element</span></span> |  <span data-ttu-id="5919d-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5919d-110">Required</span></span>  |  <span data-ttu-id="5919d-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="5919d-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5919d-112">Group</span><span class="sxs-lookup"><span data-stu-id="5919d-112">Group</span></span>](group.md)      | <span data-ttu-id="5919d-113">Sim</span><span class="sxs-lookup"><span data-stu-id="5919d-113">Yes</span></span> |  <span data-ttu-id="5919d-114">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="5919d-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="5919d-115">Rótulo</span><span class="sxs-lookup"><span data-stu-id="5919d-115">Label</span></span>](#label-tab)      | <span data-ttu-id="5919d-116">Sim</span><span class="sxs-lookup"><span data-stu-id="5919d-116">Yes</span></span> |  <span data-ttu-id="5919d-117">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="5919d-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="5919d-118">Control</span><span class="sxs-lookup"><span data-stu-id="5919d-118">Control</span></span>](control.md)    | <span data-ttu-id="5919d-119">Sim</span><span class="sxs-lookup"><span data-stu-id="5919d-119">Yes</span></span> |  <span data-ttu-id="5919d-120">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="5919d-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="5919d-121">Group</span><span class="sxs-lookup"><span data-stu-id="5919d-121">Group</span></span>

<span data-ttu-id="5919d-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="5919d-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="5919d-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="5919d-124">Label (Tab)</span></span>

<span data-ttu-id="5919d-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="5919d-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="5919d-127">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="5919d-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```