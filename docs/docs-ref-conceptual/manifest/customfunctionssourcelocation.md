# <a name="sourcelocation-element"></a><span data-ttu-id="e2e24-101">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e2e24-101">SourceLocation element</span></span>

<span data-ttu-id="e2e24-102">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="e2e24-102">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e2e24-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="e2e24-103">Attributes</span></span>

| <span data-ttu-id="e2e24-104">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="e2e24-104">**Attribute**</span></span> | <span data-ttu-id="e2e24-105">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="e2e24-105">**Required**</span></span> | <span data-ttu-id="e2e24-106">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="e2e24-106">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="e2e24-107">resid</span><span class="sxs-lookup"><span data-stu-id="e2e24-107">resid</span></span>         | <span data-ttu-id="e2e24-108">Sim</span><span class="sxs-lookup"><span data-stu-id="e2e24-108">Yes</span></span>          | <span data-ttu-id="e2e24-109">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="e2e24-109">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="e2e24-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e2e24-110">Child elements</span></span>

<span data-ttu-id="e2e24-111">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="e2e24-111">None</span></span>

## <a name="example"></a><span data-ttu-id="e2e24-112">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e2e24-112">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```