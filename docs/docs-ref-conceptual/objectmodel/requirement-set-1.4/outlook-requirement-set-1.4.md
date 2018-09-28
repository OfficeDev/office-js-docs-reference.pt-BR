# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="8d928-101">Conjunto de requisitos de API para suplementos do Outlook versão 1.4</span><span class="sxs-lookup"><span data-stu-id="8d928-101">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="8d928-102">O Outlook suplemento subconjunto de API da API do JavaScript para Office inclui objetos, métodos, propriedades e suplemento de eventos que você pode usar no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8d928-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8d928-103">Esta documentação é para um [requisito definido](/javascript/office/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisito mais recente.</span><span class="sxs-lookup"><span data-stu-id="8d928-103">This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="8d928-104">Novidades na versão 1.4</span><span class="sxs-lookup"><span data-stu-id="8d928-104">What's new in 1.4?</span></span>

<span data-ttu-id="8d928-p101">O conjunto de requisitos 1.4 inclui todos os recursos do [Conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="8d928-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="8d928-107">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="8d928-107">Change log</span></span>

- <span data-ttu-id="8d928-108">Foi adicionado o [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Exibe uma caixa de diálogo em um host do Office.</span><span class="sxs-lookup"><span data-stu-id="8d928-108">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="8d928-109">Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.</span><span class="sxs-lookup"><span data-stu-id="8d928-109">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="8d928-110">Objeto de [diálogo](/javascript/api/office/office.dialog) adicionado: O objeto retornado quando o [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) método for chamado.</span><span class="sxs-lookup"><span data-stu-id="8d928-110">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="8d928-111">Confira também</span><span class="sxs-lookup"><span data-stu-id="8d928-111">See also</span></span>

- <span data-ttu-id="8d928-112">
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)</span><span class="sxs-lookup"><span data-stu-id="8d928-112">[Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)</span></span>
- [<span data-ttu-id="8d928-113">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d928-113">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="8d928-114">Introdução</span><span class="sxs-lookup"><span data-stu-id="8d928-114">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)