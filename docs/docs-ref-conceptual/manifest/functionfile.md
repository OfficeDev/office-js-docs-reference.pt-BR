# <a name="functionfile-element"></a><span data-ttu-id="fdc51-101">Elemento FunctionFile</span><span class="sxs-lookup"><span data-stu-id="fdc51-101">FunctionFile element</span></span>

<span data-ttu-id="fdc51-p101">Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário. O elemento **FunctionFile** é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md). O atributo **resid** do elemento **FunctionFile** está definido como o valor do atributo **id** de um elemento **Url** no elemento **Resources**, que contém a URL para um arquivo HTML que armazena ou carrega todas as funções JavaScript usadas por botões de comando de suplemento sem interface de usuário, conforme definido pelo [Control element](control.md).</span><span class="sxs-lookup"><span data-stu-id="fdc51-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="fdc51-105">Veja a seguir um exemplo do elemento **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="fdc51-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="fdc51-106">O JavaScript no arquivo HTML indicado pelo elemento **FunctionFile** deve chamar `Office.initialize` e definir funções nomeadas que utilizem um único parâmetro: `event`.</span><span class="sxs-lookup"><span data-stu-id="fdc51-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="fdc51-107">As funções devem usar o `item.notificationMessages` API para indicar o progresso, êxito ou Falha ao usuário.</span><span class="sxs-lookup"><span data-stu-id="fdc51-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="fdc51-108">Ele também deve chamar `event.completed` quando concluiu a execução.</span><span class="sxs-lookup"><span data-stu-id="fdc51-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="fdc51-109">O nome das funções são usados no elemento **nomedafunção** para botões de interface do usuário sem.</span><span class="sxs-lookup"><span data-stu-id="fdc51-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="fdc51-110">Veja a seguir um exemplo de um arquivo HTML que define uma função **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="fdc51-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="fdc51-111">O código a seguir mostra como implementar a função usada pelo **nomedafunção**.</span><span class="sxs-lookup"><span data-stu-id="fdc51-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="fdc51-112">A chamada para **event.completed** sinaliza que você tenha manipulado com êxito o evento.</span><span class="sxs-lookup"><span data-stu-id="fdc51-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="fdc51-113">Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="fdc51-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="fdc51-114">O primeiro evento é executado automaticamente, enquanto os outros permanecem na fila.</span><span class="sxs-lookup"><span data-stu-id="fdc51-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="fdc51-115">Quando sua função chama **event.completed**, a próxima chamada em fila para essa função é executada.</span><span class="sxs-lookup"><span data-stu-id="fdc51-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="fdc51-116">Você deve chamar **event.completed**; Caso contrário, sua função não será executado.</span><span class="sxs-lookup"><span data-stu-id="fdc51-116">You must call **event.completed**; otherwise your function will not run.</span></span>