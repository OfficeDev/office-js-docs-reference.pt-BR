# <a name="requestedheight-element"></a><span data-ttu-id="9873c-101">Elemento RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="9873c-101">RequestedHeight element</span></span>

<span data-ttu-id="9873c-102">Especifica a altura inicial (em pixels) de um conteúdo suplemento ou o suplemento de email.</span><span class="sxs-lookup"><span data-stu-id="9873c-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="9873c-103">**Tipo de suplemento:** Conteúdo de email</span><span class="sxs-lookup"><span data-stu-id="9873c-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9873c-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="9873c-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="9873c-105">Contidos em</span><span class="sxs-lookup"><span data-stu-id="9873c-105">Contained in</span></span>

- <span data-ttu-id="9873c-106">[DefaultSettings](defaultsettings.md) (Complementos de conteúdo) com um valor que pode ser entre 32 e 1000</span><span class="sxs-lookup"><span data-stu-id="9873c-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="9873c-107">[DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (Mail suplementos) com um valor que pode ser entre 32 e 450</span><span class="sxs-lookup"><span data-stu-id="9873c-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="9873c-108">[ExtensionPoint](extensionpoint.md) (Mail contextual suplementos) com um valor que pode ser entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="9873c-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>