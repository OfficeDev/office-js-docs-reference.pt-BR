---
title: Excel conjunto de requisitos somente da API JavaScript online
description: Detalhes sobre o conjunto de requisitos do ExcelApiOnline.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ade48cec3e1db452206a1d35291b3f6f3e609566
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484829"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel conjunto de requisitos somente da API JavaScript online

O `ExcelApiOnline` conjunto de requisitos é um conjunto de requisitos especial que inclui recursos que estão disponíveis apenas para Excel na Web. ApIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações comportamentais ou estruturais não documentados) para o aplicativo Excel na Web. `ExcelApiOnline`As APIs são consideradas APIs de "visualização" para outras plataformas (Windows, Mac, iOS) e podem não ser suportadas por nenhuma dessas plataformas.

Quando as APIs no conjunto `ExcelApiOnline` de requisitos são suportadas em todas as plataformas, elas serão adicionadas ao próximo conjunto de requisitos lançado (`ExcelApi 1.[NEXT]`). Depois que esse novo requisito for público, essas APIs serão removidas de `ExcelApiOnline`. Pense nisso como um processo de promoção semelhante a uma API que está mudando da visualização para a versão.

> [!IMPORTANT]
> `ExcelApiOnline` é um superconjunto do conjunto de requisitos numerado mais recente.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` é a única versão das APIs somente online. Isso porque Excel na Web sempre terá uma única versão disponível para os usuários que são a versão mais recente.

A tabela a seguir fornece um resumo conciso das APIs, enquanto a `ExcelApiOnline` tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada das APIs atuais.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Guias de trabalho vinculadas | Gerencie links entre as guias de trabalho, incluindo o suporte para atualizar e quebrar links de livros de trabalho. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Exibições de planilha nomeadas | Fornece controle programático de exibições de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| Eventos de movimentação de planilha | Detectar quando as planilhas são movidas dentro de uma coleção, a posição da planilha e a origem da alteração. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>Uso recomendado

Como `ExcelApiOnline` as APIs só têm suporte Excel na Web, o seu complemento deve verificar se o conjunto de requisitos é suportado antes de chamar essas APIs. Isso evita chamar uma API somente online em uma plataforma diferente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Depois que a API está em um conjunto de requisitos entre plataformas, você deve remover ou editar a `isSetSupported` verificação. Isso habilita o recurso do seu complemento em outras plataformas. Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.

> [!IMPORTANT]
> Seu manifesto não pode especificar `ExcelApiOnline 1.1` como um requisito de ativação. Não é um valor válido a ser usado no [elemento Set](/javascript/api/manifest/set).

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente incluídas no conjunto `ExcelApiOnline` de requisitos. Para uma lista completa de todas as EXCEL JavaScript (`ExcelApiOnline`incluindo APIs e APIs lançadas anteriormente), consulte todas as [APIs JavaScript Excel JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-online.md)]
## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [APIs de visualização do JavaScript para Excel](excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
