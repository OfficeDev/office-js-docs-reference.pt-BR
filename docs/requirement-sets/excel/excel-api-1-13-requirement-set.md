---
title: Excel conjunto de requisitos da API JavaScript 1.13
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 79e493e25fd5c0a72cf62e0b1af9ce06fb262055
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484922"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Novidades na API JavaScript 1.13 Excel JavaScript

O ExcelApi 1.13 adicionou um método para inserir planilhas em uma pasta de trabalho de uma cadeia de caracteres codificada com Base64 e um evento para detectar a ativação de pasta de trabalho. Ele também aumentou o suporte a fórmulas em intervalos adicionando APIs para rastrear alterações nas fórmulas e localizar células dependentes diretas de uma fórmula. Além disso, ele expandiu o suporte à Tabela Dinâmica adicionando APIs pivotLayout para o gerenciamento de células de alt text, style e empty.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Fórmula de eventos alterados](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#detect-formula-changes) | Acompanhe as alterações nas fórmulas, incluindo a origem e o tipo de evento que causou uma alteração. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Dependentes da fórmula](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Localize as células dependentes diretas de uma fórmula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Inserir planilhas](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insira planilhas de outra pasta de trabalho na pasta de trabalho atual como uma cadeia de caracteres codificada com Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotLayout de tabela dinâmica](/office/dev/add-ins/excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Uma expansão da classe PivotLayout, incluindo novo suporte para alt text e gerenciamento de células vazias. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.13. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.13 ou anterior, consulte Excel APIs no conjunto de requisitos [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true) ou anterior.

[!INCLUDE[API table](../../includes/excel-1_13.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
