# <a name="metadata-element"></a><span data-ttu-id="f6189-101">Elemento de metadados</span><span class="sxs-lookup"><span data-stu-id="f6189-101">Metadata element</span></span>

<span data-ttu-id="f6189-102">Define as configurações de metadados usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="f6189-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f6189-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6189-103">Attributes</span></span>

<span data-ttu-id="f6189-104">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="f6189-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="f6189-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f6189-105">Child elements</span></span>

|  <span data-ttu-id="f6189-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="f6189-106">Element</span></span>  |  <span data-ttu-id="f6189-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f6189-107">Required</span></span>  |  <span data-ttu-id="f6189-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6189-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f6189-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f6189-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="f6189-110">Sim</span><span class="sxs-lookup"><span data-stu-id="f6189-110">Yes</span></span>  | <span data-ttu-id="f6189-111">Cadeia de caracteres com a id de recurso do arquivo JSON usado pelas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f6189-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="f6189-112">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6189-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
