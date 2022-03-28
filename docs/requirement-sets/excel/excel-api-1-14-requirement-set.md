---
title: Excel conjunto de requisitos da API JavaScript 1.14
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.14.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 44f3b56cf0bf9d5366cdfe7c2ae0454b03c2d3d6
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485252"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Novidades na API JavaScript 1.14 Excel JavaScript

O ExcelApi 1.14 adicionou objetos para controlar o recurso de tabela de dados de um gráfico, um método para localizar todas as células precedentes de uma fórmula e eventos de proteção de planilha para rastrear alterações no estado de proteção de uma planilha. Ele também adicionou vários [`getItemOrNullObject`](/office/dev/add-ins/develop/application-specific-api-model.md#ornullobject-methods-and-properties) métodos para objetos como `CommentCollection`, `ShapeCollection`e para melhorar `StyleCollection` o tratamento de erros.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Tabelas de dados de gráfico](/office/dev/add-ins/excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Controlar a aparência, a formatação e a visibilidade das tabelas de dados nos gráficos. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Precedentes de fórmula](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Retorne todas as células precedentes de uma fórmula. | [Range](/javascript/api/excel/excel.range) |
| Consultas | Recupere atributos de Consulta do Power, como nome, data de atualização e contagem de consultas. | [Consulta](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [Eventos de proteção de planilha](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Acompanhe as alterações no estado de proteção de uma planilha e na origem dessas alterações. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.14. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.14 ou anterior, consulte Excel APIs no conjunto de requisitos [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true) ou anterior.

[!INCLUDE[API table](../../includes/excel-1_14.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
