---
title: Excel conjunto de requisitos da API JavaScript 1.9
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: d3844285442e44820b5a9bc1e0aeebc92c336e65
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485115"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Novidades na API JavaScript 1.9 Excel JavaScript

Mais de 500 novas APIs do Excel foram introduzidas com o conjunto de requisitos 1.9. A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Shapes](/office/dev/add-ins/excel/excel-add-ins-shapes) | Inserir, posicionar e formatar imagens, formas geométricas e caixas de texto. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Filtro Automático](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#filter-data) | Adicionar filtros aos intervalos. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges) | Suporte para intervalos descontínuos. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Células Especiais](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Obtenha células que contêm datas, comentários ou fórmulas dentro de um intervalo. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](/office/dev/add-ins/excel/excel-add-ins-ranges-string-match) | Encontre valores ou fórmulas em uma planilha ou intervalo. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copiar e colar](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | Copie fórmulas, formatos e valores de um intervalo para outro. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](/office/dev/add-ins/excel/performance.md#suspend-calculation-temporarily) | Maior controle sobre o mecanismo de cálculo do Excel. | [Aplicativo](/javascript/api/excel/excel.application) |
| Novos Gráficos | Explore os novos tipos de gráficos compatíveis: mapas, caixa estreita, cascata, explosão solar, pareto. e funil. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | Novos recursos com formatos de intervalo. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.9. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.9 ou anterior, consulte Excel APIs no conjunto de requisitos [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) ou anterior.

[!INCLUDE[API table](../../includes/excel-1_9.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
