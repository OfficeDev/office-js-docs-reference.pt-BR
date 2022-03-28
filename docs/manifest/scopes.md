---
title: Elemento Scopes no arquivo de manifesto
description: O elemento Scopes contém permissões que o add-in precisa para se conectar a um recurso externo.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3b67977421a6ed0da205a051d0ecaf480c96fd5c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484873"
---
# <a name="scopes-element"></a>Elemento Scopes

Contém permissões que o complemento precisa para um recurso externo, como o Microsoft Graph. Quando o microsoft Graph é o recurso, o AppSource usa o elemento Scopes para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.

**Tipo de complemento:** Painel de tarefas, Email, Conteúdo

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Conteúdo 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [IdentityAPI 1.3](../requirement-sets/common/identity-api-requirement-sets.md)

**Escopos** é um elemento filho do [elemento WebApplicationInfo](webapplicationinfo.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Escopo**                |  Sim     |   O nome de uma permissão; por exemplo, Files.Read.All ou profile. |

## <a name="example"></a>Exemplo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
