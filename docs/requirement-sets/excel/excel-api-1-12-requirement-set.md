---
title: Excel conjunto de requisitos da API JavaScript 1.12
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 01cf9523f60fbf374108a81de841840310a4a875
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484923"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Novidades na API JavaScript 1.12 Excel JavaScript

O ExcelApi 1.12 aumentou o suporte a fórmulas em intervalos adicionando APIs para controlar matrizes dinâmicas e encontrando precedentes diretos de uma fórmula. Ele também adicionou controle API de filtros de tabela dinâmica. Melhorias também foram feitas nas áreas de recurso comentário, configurações de cultura e propriedades personalizadas.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Eventos de comentário](/office/dev/add-ins/excel/excel-add-ins-comments.md#comment-events) | Adiciona eventos para adicionar, alterar e excluir à coleção de comentários.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Configurações de cultura [de data e hora](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Dá acesso a configurações culturais adicionais em torno da formatação de data e hora. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Aplicativo NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Precedentes diretos](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents) | Retorna intervalos usados para avaliar a fórmula de uma célula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Filtros pivôs | Aplica filtros orientados por valor aos campos de uma tabela dinâmica. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Vazamento de intervalo](/office/dev/add-ins/excel/excel-add-ins-ranges-dynamic-arrays) | Permite que os complementos encontrem intervalos associados aos [resultados dinâmicos da matriz](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Propriedades personalizadas no nível da planilha](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Permite que as propriedades personalizadas sejam escopo para o nível da planilha, além de serem escopo para o nível da pasta de trabalho. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.12. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.12 ou anterior, consulte Excel APIs no conjunto de requisitos [1.12 ou anterior](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_12.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
