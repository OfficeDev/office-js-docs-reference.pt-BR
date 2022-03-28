---
title: Excel conjunto de requisitos da API JavaScript 1.7
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1dc99b6dbacf4a7ffc0e3672155bbde175da3698
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485247"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

## <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. A seguir estão alguns exemplos.

- Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
- Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
- Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
- Legenda do gráfico - formate a fonte de legenda de um gráfico.
- Ponto do gráfico - defina a cor do ponto do gráfico.
- Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
- Tipo de gráfico - opção para criar mais tipos de gráfico.

## <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](/office/dev/add-ins/excel/excel-add-ins-events).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

- Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
- Modificar a cor da guia de planilha.
- Adicione títulos de planilha.

Você pode personalizar a aparência de intervalos de várias maneiras:

- Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
- Defina a orientação de texto para um intervalo.
- Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

## <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

## <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

## <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

- Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
- Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
- Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no Excel de requisitos da API JavaScript 1.7. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.7 ou anterior, consulte Excel APIs no conjunto de requisitos [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true) ou anterior.

[!INCLUDE[API table](../../includes/excel-1_7.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
