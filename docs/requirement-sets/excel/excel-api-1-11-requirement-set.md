---
title: Excel conjunto de requisitos da API JavaScript 1.11
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1d002e1884e8b18981b00160d8a498ec0ff3ea51
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503180"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Novidades na EXCEL JavaScript 1.11

O ExcelApi 1.11 aprimorou o suporte para comentários e controles no nível da planilha (como salvar e fechar a planilha). Ele também adicionou acesso às configurações de cultura para ajudar a contabilizar a localização.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Menções de comentário](/office/dev/add-ins/excel/excel-add-ins-comments#mentions) |Marca e notifica outros usuários da área de trabalho por meio de comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Resolução de [comentários](/office/dev/add-ins/excel/excel-add-ins-comments#resolve-comment-threads) | Resolver threads de comentário e obter o status da resolução. | [Comment](/javascript/api/excel/excel.comment) |
| [Configurações de cultura](/office/dev/add-ins/excel/excel-add-ins-workbooks#access-application-culture-settings) | Obtém configurações do sistema cultural para a caixa de trabalho, como formatação de número. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Aplicativo NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Cortar e colar (moveTo)](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | Replica a funcionalidade de recorte e colar no Excel para um Range. | [Range](/javascript/api/excel/excel.range) |
| [Salvar](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#save-the-workbook) e [Fechar](/office/dev/add-ins/excel/excel-add-ins-workbooks#close-the-workbook) a pasta de trabalho | Salve e feche a pasta de trabalho. | [Workbook](/javascript/api/excel/excel.workbook) |
| Eventos de planilha | Eventos adicionais e informações de evento para cálculos de planilha e linhas ocultas. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.11. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.11 ou anterior, consulte Excel APIs no conjunto de requisitos [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true) ou anterior.

[!INCLUDE[API table](../../includes/excel-1_11.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
