# <a name="desktopformfactor-element"></a><span data-ttu-id="f8d0d-101">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f8d0d-101">DesktopFormFactor element</span></span>

<span data-ttu-id="f8d0d-p101">Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office para Windows, Office para Mac e Office Online. Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="f8d0d-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="f8d0d-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="f8d0d-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="f8d0d-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f8d0d-107">Child elements</span></span>

| <span data-ttu-id="f8d0d-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="f8d0d-108">Element</span></span>                               | <span data-ttu-id="f8d0d-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f8d0d-109">Required</span></span> | <span data-ttu-id="f8d0d-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="f8d0d-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="f8d0d-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="f8d0d-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="f8d0d-112">Sim</span><span class="sxs-lookup"><span data-stu-id="f8d0d-112">Yes</span></span>      | <span data-ttu-id="f8d0d-113">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="f8d0d-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="f8d0d-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="f8d0d-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="f8d0d-115">Sim</span><span class="sxs-lookup"><span data-stu-id="f8d0d-115">Yes</span></span>      | <span data-ttu-id="f8d0d-116">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f8d0d-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="f8d0d-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="f8d0d-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="f8d0d-118">Não</span><span class="sxs-lookup"><span data-stu-id="f8d0d-118">No</span></span>       | <span data-ttu-id="f8d0d-119">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f8d0d-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="f8d0d-120">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f8d0d-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
