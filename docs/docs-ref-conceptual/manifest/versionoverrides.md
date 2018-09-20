# <a name="versionoverrides-element"></a><span data-ttu-id="5a702-101">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5a702-101">VersionOverrides element</span></span>

<span data-ttu-id="5a702-p101">O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="5a702-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="5a702-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="5a702-105">Attributes</span></span>

|  <span data-ttu-id="5a702-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="5a702-106">Attribute</span></span>  |  <span data-ttu-id="5a702-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5a702-107">Required</span></span>  |  <span data-ttu-id="5a702-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a702-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5a702-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="5a702-109">**xmlns**</span></span>       |  <span data-ttu-id="5a702-110">Sim</span><span class="sxs-lookup"><span data-stu-id="5a702-110">Yes</span></span>  |  <span data-ttu-id="5a702-111">O local do esquema, que deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` quando `xsi:type` for `VersionOverridesV1_0` e `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` quando `xsi:type` for `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="5a702-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="5a702-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="5a702-112">**xsi:type**</span></span>  |  <span data-ttu-id="5a702-113">Sim</span><span class="sxs-lookup"><span data-stu-id="5a702-113">Yes</span></span>  | <span data-ttu-id="5a702-p102">A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="5a702-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="5a702-116">Atualmente, somente 2016 Outlook suporta o esquema do VersionOverrides v 1.1 e o `VersionOverridesV1_1` tipo.</span><span class="sxs-lookup"><span data-stu-id="5a702-116">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5a702-117">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="5a702-117">Child elements</span></span>

|  <span data-ttu-id="5a702-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="5a702-118">Element</span></span> |  <span data-ttu-id="5a702-119">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5a702-119">Required</span></span>  |  <span data-ttu-id="5a702-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a702-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5a702-121">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="5a702-121">**Description**</span></span>    |  <span data-ttu-id="5a702-122">Não</span><span class="sxs-lookup"><span data-stu-id="5a702-122">No</span></span>   |  <span data-ttu-id="5a702-p103">Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](./resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.</span><span class="sxs-lookup"><span data-stu-id="5a702-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="5a702-127">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="5a702-127">**Requirements**</span></span>  |  <span data-ttu-id="5a702-128">Não</span><span class="sxs-lookup"><span data-stu-id="5a702-128">No</span></span>   |  <span data-ttu-id="5a702-p104">Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5a702-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="5a702-131">Hosts</span><span class="sxs-lookup"><span data-stu-id="5a702-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="5a702-132">Sim</span><span class="sxs-lookup"><span data-stu-id="5a702-132">Yes</span></span>  |  <span data-ttu-id="5a702-p105">Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5a702-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="5a702-135">Recursos</span><span class="sxs-lookup"><span data-stu-id="5a702-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="5a702-136">Sim</span><span class="sxs-lookup"><span data-stu-id="5a702-136">Yes</span></span>  | <span data-ttu-id="5a702-137">Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5a702-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="5a702-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="5a702-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="5a702-139">Não</span><span class="sxs-lookup"><span data-stu-id="5a702-139">No</span></span>  | <span data-ttu-id="5a702-p106">Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="5a702-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="5a702-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="5a702-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="5a702-143">Não</span><span class="sxs-lookup"><span data-stu-id="5a702-143">No</span></span>  | <span data-ttu-id="5a702-144">Especifica detalhes sobre o aplicativo Web associado do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5a702-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="5a702-145">Exemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5a702-145">VersionOverrides example</span></span>
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="5a702-146">Implementar várias versões</span><span class="sxs-lookup"><span data-stu-id="5a702-146">Implementing multiple versions</span></span>

<span data-ttu-id="5a702-p107">Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.</span><span class="sxs-lookup"><span data-stu-id="5a702-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="5a702-149">Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior.</span><span class="sxs-lookup"><span data-stu-id="5a702-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="5a702-150">O elemento filho `VersionOverrides` não herda os valores do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="5a702-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="5a702-151">Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="5a702-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
