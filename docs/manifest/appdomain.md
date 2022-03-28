---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que são usados pelo seu complemento e devem ser confiáveis por Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: ec0dec56e75d25a2ba79eed9932aa0158be45b7e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485004"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional que Office deve confiar, além do especificado no [elemento SourceLocation](sourcelocation.md). A especificação de um domínio tem esses efeitos:

- Ele permite que páginas, rotas ou outros recursos no domínio sejam abertos diretamente no painel de tarefas raiz do add-in em plataformas Office desktop. (Especificar um domínio em um **AppDomain** não é necessário para Office na Web ou para abrir um recurso em um IFrame, nem é necessário para abrir um recurso em uma caixa de diálogo aberta com a [API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins) de Diálogo.)
- Ele permite que as páginas no domínio façam chamadas Office.js API de IFrames dentro do complemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Se houver uma porta explícita para o domínio, inclua-a (por exemplo, ).`<AppDomain>https://myappdomain.com:9999</AppDomain>`).
> 3. Se um subdomínio precisar ser confiável, inclua-o (por exemplo, ).`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`). O subdomínio `mysubdomain.mydomain.com` e `mydomain.com` são domínios diferentes. Se ambos precisam ser confiáveis, ambos precisam estar em elementos **AppDomain** separados.
> 4. Listar o mesmo domínio especificado no [elemento SourceLocation](sourcelocation.md) não tem efeito e pode ser enganoso. Em particular, quando você está desenvolvendo `localhost`em , você não precisa criar um **elemento AppDomain** para `localhost`.
> 5. Não inclua segmentos de uma URL além do domínio. Por exemplo, não inclua a URL completa de uma página.
> 6. Não *coloque* uma barra de fechamento , "/", no valor.

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Para saber mais, confira o [manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).
