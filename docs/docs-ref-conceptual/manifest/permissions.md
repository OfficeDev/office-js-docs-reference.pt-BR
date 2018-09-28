# <a name="permissions-element"></a><span data-ttu-id="0b268-101">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="0b268-101">Permissions element</span></span>

<span data-ttu-id="0b268-102">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="0b268-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="0b268-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="0b268-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0b268-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="0b268-104">Syntax</span></span>

<span data-ttu-id="0b268-105">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="0b268-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="0b268-106">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="0b268-106">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="0b268-107">Contidos em</span><span class="sxs-lookup"><span data-stu-id="0b268-107">Contained in</span></span>

[<span data-ttu-id="0b268-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0b268-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="0b268-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="0b268-109">Remarks</span></span>

<span data-ttu-id="0b268-110">Para saber mais, confira [Solicitando permissões para API usadas em suplementos de conteúdo e de painel de tarefas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entendendo as permissões de suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="0b268-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
