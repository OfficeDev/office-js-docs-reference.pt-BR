# <a name="webapplicationinfo-element"></a><span data-ttu-id="de8e0-101">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="de8e0-101">WebApplicationInfo element</span></span>

<span data-ttu-id="de8e0-102">Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:</span><span class="sxs-lookup"><span data-stu-id="de8e0-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="de8e0-103">Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.</span><span class="sxs-lookup"><span data-stu-id="de8e0-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="de8e0-104">Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="de8e0-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="de8e0-105">**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="de8e0-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="de8e0-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="de8e0-106">Child elements</span></span>

|  <span data-ttu-id="de8e0-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="de8e0-107">Element</span></span> |  <span data-ttu-id="de8e0-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="de8e0-108">Required</span></span>  |  <span data-ttu-id="de8e0-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="de8e0-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de8e0-110">**Id**</span><span class="sxs-lookup"><span data-stu-id="de8e0-110">**Id**</span></span>    |  <span data-ttu-id="de8e0-111">Sim</span><span class="sxs-lookup"><span data-stu-id="de8e0-111">Yes</span></span>   |  <span data-ttu-id="de8e0-112">A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.</span><span class="sxs-lookup"><span data-stu-id="de8e0-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="de8e0-113">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="de8e0-113">**Resource**</span></span>  |  <span data-ttu-id="de8e0-114">Sim</span><span class="sxs-lookup"><span data-stu-id="de8e0-114">Yes</span></span>   |  <span data-ttu-id="de8e0-115">Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.</span><span class="sxs-lookup"><span data-stu-id="de8e0-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="de8e0-116">Escopos</span><span class="sxs-lookup"><span data-stu-id="de8e0-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="de8e0-117">Não</span><span class="sxs-lookup"><span data-stu-id="de8e0-117">No</span></span>  |  <span data-ttu-id="de8e0-118">Especifica as permissões que seu suplemento precisa para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="de8e0-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="de8e0-119">Atualmente, é necessário que o recurso do suplemento corresponde seu Host.</span><span class="sxs-lookup"><span data-stu-id="de8e0-119">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="de8e0-120">O Office não solicitará um token para um suplemento, a menos que possa provar a propriedade, e hoje isso é feito hospedando o suplemento sob o nome de domínio totalmente qualificado do recurso.</span><span class="sxs-lookup"><span data-stu-id="de8e0-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="de8e0-121">Exemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="de8e0-121">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
