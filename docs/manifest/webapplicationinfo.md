---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: Documentação de referência do elemento WebApplicationInfo para Office arquivos XML (manifesto de complementos).
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: db4ce3e94420af385c818d98eb6c4930676d075d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484838"
---
# <a name="webapplicationinfo-element"></a>Elemento WebApplicationInfo

Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:

- Um recurso OAuth *2.0 para* o qual o Office cliente pode precisar de permissões.
- Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.

**Tipo de complemento:** Painel de tarefas, Email, Conteúdo

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Conteúdo 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [IdentityAPI 1.3](../requirement-sets/common/identity-api-requirement-sets.md)

> [!NOTE]
> No momento, a API de login único tem suporte para Word, Excel, Outlook e PowerPoint. Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](../requirement-sets/common/identity-api-requirement-sets.md). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Microsoft 365. Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.  

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Id**    |  Sim   |  A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.|
|  **Recurso**  |  Sim   |  Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.|
|  [Escopos](scopes.md)                |  Sim  |  Especifica as permissões que o complemento precisa para um recurso, como o Microsoft Graph.  |

## <a name="webapplicationinfo-example"></a>Exemplo de WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
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
