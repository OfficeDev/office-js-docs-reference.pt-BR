# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos de API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Suplementos do Office usam conjuntos de requisito especificados no manifesto ou uma seleção de tempo de execução para determinar se um host do Office oferece suporte a APIs que precisa de um suplemento. Para obter mais informações, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Suplementos do Excel enfrentar diversas versões do Office, incluindo o Office 2016 para Windows, o Office para iPad, o Office para Mac e Office Online. A tabela a seguir lista os conjuntos de requisito do Excel, os aplicativos de host do Office que suportam a cada conjunto de requisitos e as versões de compilação ou número para esses aplicativos.

> [!NOTE]
> Qualquer API que está marcado como **Beta** não está pronto para a produção de usuário final. Podemos disponibilizá-los para que os desenvolvedores tente-los a saída em ambientes de teste e desenvolvimento. Eles não devem ser usadas contra documentos críticos de produção de negócios.
> 
> Para os conjuntos de requisito que são marcados como **Beta**, use a versão especificada (ou posterior) do software do Office e usar a biblioteca Beta na CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entradas não marcadas como **Beta** geralmente estão disponíveis e você pode usar a biblioteca de produção na CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Conjunto de requisitos  |  Office 365 para Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | Por favor, [visite nossa página de especificação open API do JavaScript Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.7  | Versão 1801 (Build 9001.2171) ou posterior| 2,9 ou posterior | 16,9 ou posterior | Abril de 2018 | Em breve |
| ExcelApi1.6  | Versão 1704 (Compilação 8201.2001) ou posterior| 2.2 ou posterior |15.36 ou posterior| Abril de 2017 | Em breve|
| ExcelApi1.5  | Versão 1703 (Compilação 8067.2070) ou posterior| 2.2 ou posterior |15.36 ou posterior| Março de 2017 | Em breve|
| ExcelApi1.4 | Versão 1701 (build 7870.2024) ou posterior| 2.2 ou posterior |15.36 ou posterior| Janeiro de 2017 | Em breve|
| ExcelApi1.3  | Versão 1608 (build 7369.2055) ou posterior| 1.27 ou posterior |  15.27 ou posterior| Setembro de 2016 | Versão 1608 (build 7601.6800) ou posterior|
| ExcelApi1.2  | Versão 1601 (build 6741.2088) ou posterior | 1.21 ou posterior | 15.22 ou posterior| janeiro de 2016 ||
| ExcelApi1.1  | Versão 1509 (build 4266.1001) ou posterior | 1.19 ou posterior | 15.20 ou posterior| janeiro de 2016 ||

> [!NOTE]
> O número de compilação para 2016 Office instalado por meio do MSI é 16.0.4266.1001. Esta versão contém apenas o conjunto de requisito ExcelApi 1.1.

Para obter mais informações sobre versões, números de compilação e servidor on-line do Office, consulte:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-17"></a>O que há de novo no Excel JavaScript API 1.7

Os recursos do Excel JavaScript API requisito set 1.7 incluem APIs para gráficos, eventos, validação de dados, planilhas, intervalos, propriedades de documento, denominado itens, opções de proteção e estilos.

### <a name="customize-charts"></a>Personalizar gráficos

Com o novo gráfico APIs, você pode criar tipos de gráfico adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título do eixo, adicionar a unidade de exibição, adiciona uma linha de tendência com média móvel, alterar uma linha de tendência linear e muito mais. Eis alguns exemplos:

* Eixo de gráfico - obter, definir, formatar e remover a unidade de eixo, rótulo e título em um gráfico.
* Série de gráficos - adicionar, definir e excluir uma série em um gráfico.  Alterar os marcadores de série, as ordens de plotagem e dimensionamento.
* Gráfico de linhas de tendência - adicionar, obter e formatar as linhas de tendência em um gráfico.
* Legenda de gráfico - formato de fonte da legenda em um gráfico.
* Ponto de gráfico - Definir cor do ponto de gráfico.
* Gráfico subcadeia de caracteres do título - obter e definir subcadeia de caracteres do título de um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

### <a name="events"></a>Eventos

Eventos no Excel APIs fornecem uma variedade de manipuladores de eventos que permitem o add-in executar automaticamente uma função designada quando um evento específico ocorre. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão atualmente disponíveis, consulte [trabalhar com eventos usando a API do JavaScript Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência das planilhas e intervalos

Usando as novas APIs, você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter colunas ou linhas específicas visíveis quando você rolar na planilha. Por exemplo, se a primeira linha em sua planilha contiver cabeçalhos, você pode congelar essa linha para que os cabeçalhos de coluna permanecerá visíveis enquanto você rola para baixo na planilha.
* Modificar a cor da guia de planilha.
* Adicione títulos da planilha.


Você pode personalizar a aparência dos intervalos de várias maneiras:

* Definir o estilo de célula para um intervalo verificar se todas as células no intervalo possuem formatação consistente. Um estilo de célula é um conjunto definido de formatação características, como fontes e tamanhos de fonte, formatos de número, bordas de célula e sombreamento da célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação do texto para um intervalo.
* Adicionar ou modificar um hiperlink em um intervalo que os links para outro local na pasta de trabalho ou para um local externo.

### <a name="manage-document-properties"></a>Gerenciar propriedades de documento

Usando os APIs de propriedades do documento, você pode acessar as propriedades de documento internas e também criar e gerenciar propriedades de documento personalizadas para armazenar o estado da lógica de pasta de trabalho e a unidade de negócios e de fluxo de trabalho.

### <a name="copy-worksheets"></a>Planilhas de cópia

Usando a cópia de planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha dentro da mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

### <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando o intervalo de várias APIs, você pode fazer coisas como get a região redor, obtenha um intervalo redimensionado e muito mais. Essas APIs deve fazer tarefas como manipulação de intervalo e lidando com muito mais eficiente.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use essas APIs para proteger os dados em uma planilha e a estrutura de pasta de trabalho.
* Atualizar um item nomeado - use essa API para atualizar um item nomeado.
* Obtenha a célula ativa - use essa API para obter a célula ativa de uma pasta de trabalho.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > chartType|Representa o tipo do gráfico. Os valores possíveis são: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|A id exclusiva do gráfico. Somente leitura.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > showAllFieldButtons|Representa se deseja exibir todos os botões de campos em um gráfico dinâmico.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relacionamento_ > borda|Representa o formato de borda da área do gráfico, que inclui a cor, linestyle e espessura. Somente leitura.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Método_ > getItem (tipo: string, agrupar: cadeia de caracteres)|Retorna o eixo específico identificado pelo tipo e o grupo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisBetweenCategories|Indica se o eixo dos valores cruza o eixo das categorias entre categorias.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisGroup|Representa o grupo do eixo especificado. Somente leitura. Os valores possíveis são: primária, secundário.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > categoryType|Retorna ou define o tipo de eixo de categorias. Os valores possíveis são: automático, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > cruza|Representa o eixo especificado onde o outro eixo o atravessa. Os valores possíveis são: automática, mínimo, máximo personalizado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > crossesAt|Representa o eixo especificado onde o outro eixo cruza no. Somente leitura. Definido como essa propriedade deverá usar o método de SetCrossesAt(double). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > customDisplayUnit|Representa o valor de unidade de exibição personalizados do eixo. Somente leitura. Para definir essa propriedade, use o método SetCustomDisplayUnit(double). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > displayUnit|Representa a unidade de exibição do eixo. Os valores possíveis são: None, centenas, milhares, TenThousands, HundredThousands, milhões, TenMillions, HundredMillions, bilhões, trilhões, personalizado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > height|Representa a altura, em pontos, do eixo de gráfico. Nulo se o eixo não do visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > esquerda|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não do visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > logBase|Representa a base do logaritmo ao usar a escala logarítmica.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > reversePlotOrder|Indica se o Microsoft Excel plota os pontos de dados do último ao primeiro.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > scaleType|Representa o tipo de escala do eixo de valor. Os valores possíveis são: Linear, logarítmica.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > showDisplayUnitLabel|Indica se o rótulo de unidade de exibição do eixo é visível.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickLabelSpacing|Representa o número de categorias ou séries entre os rótulos de marca de escala. Pode ser um valor de 1 31999 até ou uma sequência vazia para configuração automática. O valor retornado é sempre um número.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickMarkSpacing|Representa o número de categorias ou séries entre marcas de escala.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > superior|Representa a distância, em pontos, da borda superior do eixo para o topo da área do gráfico. Nulo se o eixo não do visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tipo|Representa o tipo de eixo. Somente leitura. Os valores possíveis são: inválido, categoria, valor, série.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > visível|Um valor booleano representa a visibilidade do eixo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > width|Representa a largura, em pontos, do eixo de gráfico. Nulo se o eixo não do visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > baseTimeUnit|Retorna ou define a unidade base do eixo de categorias especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > majorTickMark|Representa o tipo de marca de escala principais do eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > majorTimeUnitScale|Retorna ou define o valor de escala de unidades principal para o eixo de categoria quando a propriedade CategoryType estiver definida como a escala de tempo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > minorTickMark|Representa o tipo de marca de escala secundárias do eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > minorTimeUnitScale|Retorna ou define o valor de escala de unidades secundária do eixo das categorias quando a propriedade CategoryType estiver definida como a escala de tempo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > tickLabelPosition|Representa a posição dos rótulos de marca de escala no eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCategoryNames(sourceData: Range)|Define todos os nomes de categoria para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCrossesAt(value: double)|Defina o eixo especificado onde o outro eixo cruza no.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCustomDisplayUnit(value: double)|Define a unidade de exibição do eixo para um valor personalizado.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > color|Código de cor HTML que representa a cor das bordas do gráfico.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > peso|Representa a espessura da borda, em pontos.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relacionamento_ > lineStyle|Representa o estilo de linha da borda.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > posição|Valor de DataLabelPosition que representa a posição do rótulo de dados. Os valores possíveis são: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > separador|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showBubbleSize|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showCategoryName|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showLegendKey|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showPercentage|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showSeriesName|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showValue|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > height|Representa a altura da legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > esquerda|Representa esquerda de uma legenda de gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > showShadow|Representa se a legenda tem sombra no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > superior|Representa a parte superior de uma legenda de gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > width|Representa a largura da legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relacionamento_ > legendEntries|Representa uma coleção de legendEntries na legenda. Somente leitura.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > visível|Representa o visível de uma entrada de legenda do gráfico.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Propriedade_ > itens|Uma coleção de objetos chartLegendEntry. Somente leitura.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getCount()|Retorna o número de legendEntry na coleção.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getItemAt(index: number)|Retorna um legendEntry no índice fornecido.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > hasDataLabel|Indica se um ponto de dados tem datalabel. Não se aplica a gráficos de superfície.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerBackgroundColor|Ponto de representação de código de cores HTML da cor de plano de fundo do marcador de dados. E.g. #FF0000 representa vermelho.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerForegroundColor|Ponto de representação de código de cores HTML da cor de primeiro plano do marcador de dados. E.g. #FF0000 representa vermelho.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerSize|Representa o tamanho do marcador de ponto de dados.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerStyle|Representa o estilo de marcador de um ponto de dados do gráfico. Os valores possíveis são: inválido, automático, nenhum, quadrado, losango, triângulo, X, por estrelas, Dot, tracejado, Circle, além, imagem.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Relacionamento_ > dataLabel|Retorna o rótulo de dados de um ponto de gráfico. Somente leitura.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relacionamento_ > borda|Representa o formato de borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e espessura. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > chartType|Representa o tipo de gráfico de uma série. Os valores possíveis são: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > doughnutHoleSize|Representa o tamanho do orifício de rosca uma série de gráficos.  Válido somente em gráficos de rosca e doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > filtrada|Valor booleano que representa se a série for filtrada ou não. Não se aplica a gráficos de superfície.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > gapWidth|Representa a largura do intervalo de uma série de gráfico.  Válido apenas diante de barra e gráficos de coluna, bem como|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > hasDataLabels|Valor booleano que representa se a série tiver rótulos de dados ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerBackgroundColor|Representa a cor de plano de fundo de marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerForegroundColor|Representa a cor de primeiro plano de marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerSize|Representa o tamanho do marcador de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerStyle|Representa o estilo de marcador de uma série de gráfico. Os valores possíveis são: inválido, automático, nenhum, quadrado, losango, triângulo, X, por estrelas, Dot, tracejado, Circle, além, imagem.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > plotOrder|Representa a ordem de plotagem de uma série de gráfico dentro do grupo gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > showShadow|Valor booleano que representa se a série tiver sombra ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > suave|Valor booleano que representa se a série é suave ou não. Somente para gráficos de linhas e de dispersão.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relacionamento_ > dataLabels|Representa uma coleção de todos os dataLabels na série. Somente leitura.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relacionamento_ > trendlines|Representa uma coleção de linhas de tendência na série. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > Delete)|Exclui a série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setBubbleSizes(sourceData: Range)|Defina os tamanhos de bolha para uma série de gráfico. Funciona somente para gráficos de bolha.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setValues(sourceData: Range)|Defina os valores de uma série de gráfico. Gráfico de dispersão, isso significa que valores do eixo Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setXAxisValues(sourceData: Range)|Defina os valores do eixo para uma série de gráfico X. Funciona somente para gráficos de dispersão.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > Adicionar (nome: cadeia de caracteres, de índice: número)|Adicione uma nova série à coleção.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > height|Retorna a altura, em pontos, do título do gráfico. Somente leitura. Nulo se o título de gráfico não do visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para o título do gráfico. Os valores possíveis são: centro, esquerda, Justify, distribuído, à direita.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > esquerda|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. Nulo se o título de gráfico não do visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > posição|Representa a posição do título do gráfico. Os valores possíveis são: Top, automático, inferior, esquerda, direita.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > showShadow|Representa um valor boolean que determina se o título do gráfico tem uma sombra.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > textOrientation|Representa a orientação do texto do título do gráfico. O valor deve ser um inteiro de – 90 a 90 ou 180 para texto orientado verticalmente.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > superior|Representa a distância, em pontos, da borda superior do título do gráfico para o topo da área do gráfico. Nulo se o título de gráfico não do visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical do título do gráfico. Os valores possíveis são: centro, inferior, Top, Justify, distribuído.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > width|Retorna a largura, em pontos, do título do gráfico. Somente leitura. Nulo se o título de gráfico não do visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Método_ > setFormula(formula: string)|Define um valor string que representa a fórmula do título do gráfico usando a notação de estilo A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relacionamento_ > borda|Representa o formato de borda do título do gráfico, que inclui a cor, linestyle e espessura. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > com versões anteriores|Representa o número de períodos em que a linha de tendência se estende para trás.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayEquation|True se a equação para a linha de tendência é exibida no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayRSquared|True se o quadrada da linha de tendência é exibida no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > encaminhar|Representa o número de períodos em que a linha de tendência se estende para frente.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > intercepção|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticas). O valor retornado é sempre um número.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > movingAveragePeriod|Representa o período de uma linha de tendência do gráfico, apenas para a linha de tendência com MovingAverage tipo.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > nome|Representa o nome da linha de tendência. Pode ser definido como um valor string ou pode ser definido como um valor nulo representa os valores automática. O valor retornado é sempre uma cadeia de caracteres|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > polynomialOrder|Representa a ordem de uma linha de tendência do gráfico, apenas para a linha de tendência com tipo polinomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > tipo|Representa o tipo de uma linha de tendência do gráfico. Os valores possíveis são: Linear, exponencial, logarítmica, MovingAverage, polinomial, energia.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relacionamento_ > formato|Representa a formatação de uma linha de tendência do gráfico. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Método_ > Delete)|Exclua o objeto de linha de tendência.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Propriedade_ > itens|Uma coleção de objetos chartTrendline. Somente leitura.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > add(type: string)|Adiciona uma nova linha de tendência à coleção de linha de tendência.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getCount()|Retorna o número de linhas de tendência na coleção.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getItem(index: number)|Obtenha o objeto trendline pelo índice, que é a ordem de inserção na matriz de itens.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relacionamento_ > linha|Representa a formatação de linha do gráfico. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > key|Obtém a chave da propriedade personalizada. Somente leitura. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > tipo|Obtém o tipo de valor da propriedade personalizada. Somente leitura. Somente leitura. Os valores possíveis são: número, booleano, data, cadeia de caracteres, flutuante.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > value|Obtém ou define o valor da propriedade personalizada.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Método_ > Delete)|Exclui a propriedade personalizada.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Propriedade_ > itens|Uma coleção de objetos customProperty. Somente leitura.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > Adicionar (chave: cadeia de caracteres, valor: objeto)|Cria uma nova propriedade personalizada ou define uma existente.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > deleteAll()|Exclui todas as propriedades personalizadas nesta coleção|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getCount()|Obtém a contagem das propriedades personalizadas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItem(key: string)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Gera se a propriedade personalizada não existe.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto nulo se a propriedade personalizada não existe.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Propriedade_ > itens|Uma coleção de objetos dataConnection. Somente leitura.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Método_ > refreshAll()|Atualiza todas as conexões de dados na coleção.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > author|Obtém ou define o autor da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > category|Obtém ou define a categoria da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > comments|Obtém ou define os comentários da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > company|Obtém ou define a empresa da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > keywords|Obtém ou define as palavras-chave da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > lastAuthor|Obtém o último autor da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > manager|Obtém ou define o gerente da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > revisionNumber|Obtém o número de revisão da pasta de trabalho. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > subject|Obtém ou define o assunto da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > title|Obtém ou define o título da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relação_ > creationDate|Obtém a data de criação da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relacionamento_ > personalizado|Obtém a coleção das propriedades personalizadas da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > fórmula|Obtém ou define a fórmula do item nomeado.  Fórmula sempre começa com um sinal de '='.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relacionamento_ > arrayValues|Retorna um objeto que contém valores e os tipos de item nomeado. Somente leitura.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > tipos|Representa os tipos de cada item na matriz item nomeado somente leitura. Os valores possíveis são: desconhecido, vazio, cadeia de caracteres, número inteiro, Double, Boolean, erro.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > values|Representa os valores de cada item na matriz item nomeado. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireColumn|Representa se o intervalo atual for uma coluna inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireRow|Representa se o intervalo atual for uma linha inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > numberFormatLocal|Representa o código de formato número do Excel para o intervalo especificado como uma cadeia de caracteres no idioma do usuário.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > style|Representa o estilo do intervalo atual. Isso retornará null ou uma cadeia de caracteres.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getAbsoluteResizedRange (numRows: number, numColumns: número)|Obtém um objeto Range com a mesma célula superior esquerda como o objeto de intervalo atual, mas com o número especificado de linhas e colunas.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > GetImage)|O intervalo é renderizada como uma imagem codificado na base64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getSurroundingRegion()|Retorna um objeto Range que representa a região ao redor da célula superior esquerda nesse intervalo. Uma região redor é um intervalo limitado por qualquer combinação de linhas em branco e colunas em branco em relação esse intervalo.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > showCard()|Exibe o cartão para uma célula ativa se ela tiver o conteúdo de rich valor.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > textOrientation|Obtém ou define a orientação do texto de todas as células dentro do intervalo.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardHeight|Determina se a altura da linha do objeto Range equivale a altura padrão da planilha.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardWidth|Determina se o columnwidth do objeto Range equivale a largura padrão da planilha.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > address|Representa o destino de url para o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > documento..|Representa o documento.. destino do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > dica de tela|Representa a cadeia de caracteres exibida ao passar o mouse sobre o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > textToDisplay|Representa a cadeia de caracteres que é exibida na parte superior esquerda da maioria dos célula do intervalo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > addIndent|Indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula estiver definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > Recuar automaticamente|Indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula estiver definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > builtIn|Indica se o estilo for um estilo interno. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > formulaHidden|Indica se a fórmula ficará oculta quando a planilha for protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para o estilo. Os valores possíveis são: geral, esquerda, centro, à direita, preenchimento, Justify, CenterAcrossSelection, distribuído.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeAlignment|Indica se o estilo inclui as propriedades Recuar automaticamente, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel e TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeBorder|Indica se o estilo inclui as propriedades de cor, ColorIndex, LineStyle e espessura da borda.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeFont|Indica se o estilo inclui as propriedades de fonte do plano de fundo, negrito, cor, ColorIndex, FontStyle, itálico, nome, tamanho, tachado, subscrito, sobrescrito e sublinhado.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeNumber|Indica se o estilo inclui a propriedade NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includePatterns|Indica se o estilo inclui as propriedades de cor, ColorIndex, InvertIfNegative, padrão, PatternColor e PatternColorIndex interior.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeProtection|Indica se o estilo inclui as propriedades de proteção de FormulaHidden e bloqueado.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > indentLevel|Um inteiro de 0 a 250 que indica o nível de recuo para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > locked|Indica se o objeto está bloqueado quando a planilha for protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > nome|O nome do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormat|O código de formato do formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormatLocal|O código de formatação localizado do formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > orientação|A orientação do texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > OrdemDeLeitura|A ordem de leitura para o estilo. Os valores possíveis são: contexto, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > shrinkToFit|Indica se texto será automaticamente diminuído para se ajustar à largura de coluna disponível.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > textOrientation|A orientação do texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical para o estilo. Os valores possíveis são: Top, centro, inferior, Justify, distribuído.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > wrapText|Indica se o Microsoft Excel faz retorno automático do texto no objeto.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relacionamento_ > bordas|Uma coleção de borda de quatro objetos Border que representam o estilo das quatro bordas. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relacionamento_ > preenchimento|O preenchimento do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > font|Um objeto Font que representa a fonte do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Método_ > Delete)|Exclui a esse estilo.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Propriedade_ > itens|Uma coleção de objetos style. Somente leitura.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > add(name: string)]|Adiciona um novo estilo à coleção.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > getItem(name: string)|Obtém um estilo pelo nome.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > address|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de alteração que representa como o evento Changed é disparado. Os valores possíveis são: outras pessoas, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > fonte|Obtém a origem do evento. Os valores possíveis são: Local, remota.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > tableId|Obtém a id da tabela na qual os dados alterados.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha em que os dados alterados.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > isInsideTable|Indica se a seleção estiver dentro de uma tabela, o endereço será inútil se IsInsideTable for false.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > tableId|Obtém a id da tabela na qual a seleção é alterada.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha na qual a seleção é alterada.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriedade_ > nome|Obtém o nome da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relacionamento_ > conexões de dados|Atualiza todas as conexões de dados na pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > properties|Obtém as propriedades de pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > protection|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relacionamento_ > estilos|Representa uma coleção de estilos associados a pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Método_ > getActiveCell()|Obtém a célula ativa no momento da pasta de trabalho.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Propriedade_ > protected|Indica se a pasta de trabalho está protegida. Somente Leitura. Somente leitura.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > protect(password: string)|Protege uma pasta de trabalho. Falha se a pasta de trabalho foi protegida.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > unprotect(password: string)|Desprotege a uma pasta de trabalho.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > linhas de grade|Obtém ou define o sinalizador de linhas de grade da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > títulos|Obtém ou define o sinalizador de títulos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showHeadings|Obtém ou define o sinalizador de títulos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardHeight|Retorna a altura padrão (padrão) de todas as linhas na planilha, em pontos. Somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardWidth|Retorna ou define a largura padrão (padrão) de todas as colunas na planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > tabColor|Obtém ou define a cor da guia de planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relacionamento_ > freezePanes|Obtém um objeto que pode ser usado para manipular os painéis congelados na planilha somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > cópia (positionType: WorksheetPositionType, relativeTo: planilha)|Copiar uma planilha e coloca na posição especificada. Retorne a planilha copiada.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getRangeByIndexes (startRow: number, startColumn: rowCount number,: number, columnCount: número)| Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha que é ativada.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > fonte|Obtém a origem do evento. Os valores possíveis são: Local, remota.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha que é adicionada à pasta de trabalho.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > address|Obtém o endereço de intervalo que representa a área alterada de uma planilha específica.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de alteração que representa como o evento Changed é disparado. Os valores possíveis são: outras pessoas, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > fonte|Obtém a origem do evento. Os valores possíveis são: Local, remota.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha em que os dados alterados.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha é desativada.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > fonte|Obtém a origem do evento. Os valores possíveis são: Local, remota.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha que seja excluída da pasta de trabalho.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeAt (frozenRange: intervalo ou a cadeia de caracteres)|Define as células congeladas no modo de exibição de planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeColumns(count: number)|Congele a primeira coluna (s) da planilha in-loco.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeRows(count: number)|Congele as principais linhas da planilha in-loco.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocation()|Obtém um intervalo que descreve as células congeladas no modo de exibição de planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocationOrNullObject()|Obtém um intervalo que descreve as células congeladas no modo de exibição de planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > unfreeze()|Remove todos os painéis congelados na planilha.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditObjects|Representa a opção de proteção da planilha de permitir a edição de objetos.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditScenarios|Representa a opção de proteção da planilha de permitir a edição de cenários.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relacionamento_ > selectionMode|Representa a opção de proteção da planilha do modo de seleção.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > tipo|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém a identificação da planilha na qual a seleção é alterada.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>O que há de novo no Excel JavaScript API 1.6 

### <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite que os seguintes tipos de formatação condicional:

* Escala de cores
* Barra de dados
* Conjunto de ícones
* Personalizado

Além disso:

* Retorna o formato condicional é aplicado ao intervalo. 
* Remoção da formatação condicional. 
* Fornece a capacidade de prioridade e stopifTrue. 
* Obtém a coleção de toda a formatação condicional em um determinado intervalo. 
* Limpa todos os formatos condicionais ativos no intervalo atual especificado. 

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > suspendApiCalculationUntilNextSync()|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho, para garantir que todas as dependências sejam propagadas.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relacionamento_ > formato|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relacionamento_ > regra|Representa o objeto Regra neste formato condicional.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Propriedade_ > threeColorScale|Caso verdadeiro, a escala de cores terá três pontos (mínimo, médio, máximo). Caso contrário, terá dois (mínimo, máximo). Somente leitura.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relação_ > criteria|Os critérios da escala de cores. O ponto médio é opcional ao se usar uma escala de cores de dois pontos.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula1|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula2|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relacionamento_ > máximo|O critério de escala de cores de ponto máximo.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relacionamento_ > ponto médio|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relacionamento_ > mínimo|O critério de escala de cores de ponto mínimo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > color|Representação de código de cor HTML da cor da escala de cores. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > fórmula|Um número, uma fórmula ou nulo (se Type for LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > tipo|No que a fórmula condicional de ícone deve se basear. Os valores possíveis são: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > fillColor|Código de cores HTML que representa a cor de preenchimento, no formato #RRGGBB (ex.: "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveBorderColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveFillColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > fillColor|Código de cores HTML que representa a cor de preenchimento, no formato #RRGGBB (ex.: "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > gradientFill|Representação booliana para indicar se a DataBar tem um gradiente ou não.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > fórmula|A fórmula, se necessário, para avaliar a regra databar.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > tipo|O tipo de regra para databar. Os valores possíveis são: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > id|A prioridade do formato condicional dentro do ConditionalFormatCollection atual. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > prioridade|A prioridade (ou índice) dentro da coleção de formatos condicionais na qual se encontra atualmente esse formato condicional. A alteração dessa propriedade também|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > stopIfTrue|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > tipo|Um tipo de formato condicional. É possível definir somente um por vez. Somente Leitura. Somente leitura. Os valores possíveis são: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > cellValue|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > cellValueOrNullObject|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > colorScale|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > colorScaleOrNullObject|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > personalizado|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > customOrNullObject|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > dataBar|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > dataBarOrNullObject|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > iconSet|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > iconSetOrNullObject|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > predefinida|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > presetOrNullObject|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > textComparison|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > textComparisonOrNullObject|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > topBottom|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relacionamento_ > topBottomOrNullObject|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > Delete)|Exclui esse formato condicional.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRange()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRangeOrNullObject()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalFormat. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > add(type: string)|Adiciona um novo formato condicional à coleção na prioridade firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > clearAll()|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getCount()|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItem(id: string)|Retorna um formato condicional para a ID especificada.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItemAt(index: number)|Retorna um formato condicional no índice fornecido.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > fórmula|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaLocal|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaR1C1|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > fórmula|Um número ou uma fórmula, dependendo do tipo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > operator|GreaterThan ou GreaterThanOrEqual para cada tipo de regra para o formato de ícone condicional. Os valores possíveis são: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relacionamento_ > customIcon|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relação_ > type|No que a fórmula condicional de ícone deve se basear.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Propriedade_ > critério|O critério do formato condicional. Os valores possíveis são: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > color|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > id|Representa o identificador da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > sideIndex|Valor constante que indica o lado específico da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > style|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Os valores possíveis são: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > contagem|Número de objetos de borda da coleção. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalRangeBorder. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relacionamento_ > inferior|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relacionamento_ > esquerda|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relacionamento_ > direita|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relacionamento_ > superior|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItem(index: string)|Obtém um objeto de borda usando seu nome|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItemAt(index: number)|Obtém um objeto de borda usando seu índice.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Propriedade_ > color|O código de cor HTML que representa a cor do preenchimento, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Método_ > Clear)|Redefine o preenchimento.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > negrito|Representa o status da fonte em negrito.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > color|Representação de código de cor HTML para a cor do texto. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > itálico|Representa o status da fonte em itálico.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > tachado|Representa o status de tachado da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > sublinhado|Tipo de sublinhado aplicado à fonte. Os valores possíveis são: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Método_ > Clear)|Redefine os formatos de fonte.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Propriedade_ > numberFormat|Representa o código de formato de número do Excel para determinado intervalo. Desmarcado se null for passado.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relacionamento_ > bordas|Coleção de objetos de borda que se aplicam ao intervalo de formatos condicionais geral. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relacionamento_ > preenchimento|Retorna o objeto de preenchimento definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > font|Retorna o objeto de fonte definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > texto|O valor de texto do formato condicional.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > classificação|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > tipo|Formatar valores com base na classificação superior ou inferior. Os valores possíveis são: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relacionamento_ > formato|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relacionamento_ > regra|Representa o objeto Regra neste formato condicional. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisColor|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisFormat|Representação de como o eixo é determinado para uma barra de dados do Excel. Os valores possíveis são: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > barDirection|Representa a direção em que o gráfico de barras de dados deve ser baseado. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > showDataBarOnly|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relacionamento_ > lowerBoundRule|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relacionamento_ > negativeFormat|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relacionamento_ > positiveFormat|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relacionamento_ > upperBoundRule|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > reverseIconOrder|Caso verdadeiro, inverte as ordens de ícones para IconSet. Observe que não será possível definir isso se ícones personalizados forem usados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > showIconOnly|Caso verdadeiro, oculta os valores e mostra somente ícones.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > style|Caso definido, exibe a opção IconSet do formato condicional. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relação_ > criteria|Uma matriz de IconSets e critérios para as regras e os ícones personalizados potenciais para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto tipo, fórmula e operador serão ignorados quando definidos.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relacionamento_ > formato|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relacionamento_ > regra|A regra da formatação condicional.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relacionamento_ > conditionalFormats|Coleção de ConditionalFormats que formam uma interseção do intervalo. Somente leitura.|1.6|
|[range](/javascript/api/excel/excel.range)|_Método_ > calculate()|Calcula um intervalo de células em uma planilha.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relacionamento_ > formato|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relacionamento_ > regra|A regra da formatação condicional.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relacionamento_ > formato|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relacionamento_ > regra|Os critérios da formatação condicional TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relacionamento_ > internalTest|Apenas para uso interno. Somente leitura.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > calculate(markAllDirty: bool)|Calcula todas as células em uma planilha.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>O que há de novo no Excel JavaScript API 1,5

### <a name="custom-xml-part"></a>Parte XML personalizada

* Adição de uma coleção de partes XML personalizadas ao objeto workbook.
* Obter parte XML personalizada usando ID
* Obtenção de um novo conjunto com escopo de partes XML personalizadas cujos namespaces correspondam ao namespace especificado.
* Obtenha uma cadeia XML associada a uma parte.
* Forneça id e namespace de uma parte.
* Adiciona uma nova parte XML personalizada à pasta de trabalho.
* Defina a parte XML inteira.
* Exclua uma parte XML personalizada.
* Exclua um atributo com o nome especificado do elemento identificado por xpath.
* Consulte o conteúdo XML por xpath.
* Insira, atualize e exclua o atributo.

**Implementação de referência:** Consulte [aqui](https://github.com/mandren/Excel-CustomXMLPart-Demo) para conhecer uma implementação de referência que mostra como partes XML personalizadas podem ser usadas em um suplemento.

### <a name="others"></a>Outros
* `range.getSurroundingRegion()` Retorna um objeto Range que representa a região ao redor desse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.
* `getNextColumn()` e `getPreviousColumn()`, `getLast() na coluna da tabela.
* `getActiveWorksheet()` na pasta de trabalho.
* `getRange(address: string)` fora da pasta de trabalho.
* `getBoundingRange(ranges: )` Obtém o menor objeto Range que abrange os intervalos fornecidos. Por exemplo, o intervalo delimitador entre "B2:C5" e "D10:E15" é "B2:E15".
* `getCount()` em várias coleções, como itens nomeados, planilhas, tabelas etc. para obter o número de itens em uma coleção. `workbook.worksheets.getCount()`
* `getFirst()` e `getLast()` e obter o último em várias coleções, como coleções de planilhas, colunas de tabela, pontos de gráfico e exibições de intervalo.
* `getNext()` e `getPrevious()` na coleção de planilhas e colunas de tabela.
* `getRangeR1C1()` Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > id|ID da parte XML personalizada. Somente leitura.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > namespaceUri|URI do namespace da parte XML personalizada. Somente leitura.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > Delete)|Exclui a parte XML personalizada.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > getXml()|Obtém o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > setXml(xml: string)|Define o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPart. Somente leitura.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > add(xml: string)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getByNamespace(namespaceUri: string)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getCount()|Obtém o número de partes CustomXml na coleção.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPartScoped. Somente leitura.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getCount()|Obtém o número de partes CustomXML nesta coleção.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItem()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItemOrNullObject()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relacionamento_ > customXmlParts|Representa a coleção de partes XML contidas nesta pasta de trabalho. Somente leitura.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNext(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método gerará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNextOrNullObject(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método retornará um objeto nulo.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPrevious(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método gerará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPreviousOrNullObject(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método retornará um objeto nulo.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getFirst(visibleOnly: bool)|Obtém a primeira planilha na coleção.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getLast(visibleOnly: bool)|Obtém a última planilha na coleção.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4
A seguir estão que os novos acréscimos ao Excel JavaScript APIs no requisito definir 1,4.

### <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` itens com escopo de planilha ou pasta de trabalho
* `worksheet` retorna a planilha que o item nomeado tem como escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)`Adiciona um novo nome à coleção do escopo fornecido.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adiciona um novo nome à coleção do escopo fornecido usando a localidade do usuário para a fórmula.

### <a name="settings-api-in-in-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Setting](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistentes ao documento. Agora, adicionamos APIs relacionadas a configurações no namespace do Excel. Isso não oferece nova funcionalidade de rede. No entanto, assim é mais fácil manter a sintaxe de API com base em lote para reduzir a dependência em tarefas comuns relacionadas à API do Excel.

As APIs incluem `getItem()` para acessar configuração de entrada por meio da chave, `add()` para adicionar o par de configuração de chave:valor especificado na pasta de trabalho.

### <a name="others"></a>Outros

* Definir nome de coluna de tabela (a versão anterior permite somente leitura).
* Adicionar coluna de tabela ao fim da tabela (a versão anterior permite apenas em qualquer lugar, exceto o último).
* Adicione várias linhas a uma tabela de cada vez (a versão anterior só permite uma linha por vez).
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Obter item ou função de objeto null: Esta funcionalidade permite obter o objeto utilizando a chave. Se o objeto não existir, a propriedade isNullObject do objeto retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto ou não sem ter de lidar com ele por meio do tratamento de exceção. Disponível na planilha, item nomeado, associação, série de gráficos etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getCount()|Obtém o número de associações da coleção.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém um objeto binding pela ID. Se o objeto binding não existir, retornará um objeto null.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getCount()|Retorna o número de gráficos da planilha.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Método_ > getCount()|Retorna o número de pontos do gráfico da série.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > getCount()|Retorna o número de série da coleção.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > comment|Representa o comentário associado a esse nome.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > escopo|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Somente leitura. Os valores possíveis são: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > planilha|Retorna a planilha em que o item nomeado tem escopo. Gerará um erro se os itens tiverem escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > worksheetOrNullObject|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > Delete)|Exclui o nome fornecido.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > getRangeOrNullObject()|Retorna o objeto Range associado ao nome. Retornará um objeto null se o tipo do item nomeado não for um intervalo.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > Adicionar (nome: cadeia de caracteres, referência: intervalo ou a cadeia de caracteres, comentário: cadeia de caracteres)|Adiciona um novo nome à coleção do escopo fornecido.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > addFormulaLocal (nome: cadeia de caracteres, fórmula: cadeia de caracteres, comentário: cadeia de caracteres)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getCount()|Obtém o número de itens nomeados na coleção.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto getNamedItem não existir, retornará um objeto null.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getCount()|Obtém o número de tabelas dinâmicas na coleção.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNullObject(name: string)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNullObject (anotherRange: intervalo ou a cadeia de caracteres)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|Retorna o intervalo usado do objeto range determinado. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getCount()|Obtém o número de objetos RangeView na coleção.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > value|Representa o valor armazenado para esta configuração.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete)|Exclui a configuração.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > Adicionar (chave: cadeia de caracteres, valor: (qualquer))|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getCount()|Obtém o número de Configurações na coleção.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém uma entrada de configuração por meio da tecla. Se a Configuração não existir, retornará um objeto null.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settings|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getCount()]|Obtém o número de tabelas na coleção.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNullObject (chave: número ou cadeia de caracteres)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, retornará um objeto null.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getCount()|Obtém a quantidade de colunas na tabela.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNullObject (chave: número ou cadeia de caracteres)|Obtém um objeto column por nome ou ID. Se a coluna não existir, retornará um objeto null.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Método_ > getCount()|Obtém a quantidade de linhas na tabela.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > nomes|Coleção de nomes com escopo para a planilha atual. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getCount(visibleOnly: bool)|Obtém o número de planilhas na coleção.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto worksheet usando o Nome ou ID dele. Se a planilha não existir, retornará um objeto null.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.3.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Método_ > Delete)|Especifica a associação.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > Adicionar (intervalo: intervalo ou a cadeia de caracteres, bindingType: cadeia de caracteres, id: cadeia de caracteres)|Adiciona uma nova associação a um intervalo específico.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromNamedItem (nome: cadeia de caracteres, bindingType: cadeia de caracteres, id: cadeia de caracteres)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromSelection (bindingType: cadeia de caracteres, id: cadeia de caracteres)|Adiciona uma nova associação com base na seleção atual.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNull(id: string)|Obtém um objeto de associação pela ID. Se o objeto em associação não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNull(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNull(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto nameditem não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Propriedade_ > nome|Nome da Tabela Dinâmica.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > planilha|A planilha que contém a Tabela Dinâmica atual. Somente leitura.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > refresh()|Atualiza a Tabela Dinâmica.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Propriedade_ > itens|Uma coleção de objetos de Tabela Dinâmica. Somente leitura.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItem(name: string)|Obtém uma Tabela Dinâmica por nome.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNull(name: string)|Obtém uma Tabela Dinâmica por nome. Se a Tabela Dinâmica não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNull (anotherRange: intervalo ou a cadeia de caracteres)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getVisibleView()|Representa as linhas visíveis do intervalo atual.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > cellAddresses|Representa os endereços de célula da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > columnCount|Retorna o número de colunas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulas|Representa a fórmula em notação A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasLocal|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, introduced in 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > index|Retorna um valor que representa o índice da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > numberFormat|Representa o código de formato de número do Excel para determinada célula.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > rowCount|Retorna o número de linhas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > texto|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > valueTypes|Representa o tipo de dados de cada célula. Somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > values|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relação_ > rows|Representa uma coleção de exibições de tabelas associadas ao intervalo. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Método_ > getRange()|Obtém o intervalo pai associado à RangeView atual.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Propriedade_ > itens|Uma coleção de objetos rangeView. Somente leitura.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getItemAt(index: number)|Obtém uma linha de RangeView através de seu índice. Indexado com zero.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete)|Exclui a configuração.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNull(key: string)|Obtém uma entrada de configuração por meio da tecla. Se o objeto de configuração não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > defina (chave: cadeia de caracteres, valor: cadeia de caracteres)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settingCollection|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightFirstColumn|Indica se a primeira coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightLastColumn|Indica se a última coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedColumns|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedRows|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showFilterButton|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNull (chave: número ou cadeia de caracteres)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNull (chave: número ou cadeia de caracteres)|Obtém um objeto de coluna por nome ou ID. Se a coluna não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > pivotTables|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > pivotTables|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Quais são as novidades na API JavaScript do Excel 1.2

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.2.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|Obtém um gráfico com base em sua posição na coleção. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > planilha|A planilha que contém o gráfico atual. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Método_ > getImage (altura: number, largura: number, fittingMode: cadeia de caracteres)|Processa o gráfico como uma imagem codificada em base64, dimensionando o gráfico para se ajustar às dimensões especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relação_ > criteria|O filtro aplicado no momento à coluna fornecida. Somente leitura.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > apply(criteria: FilterCriteria)|Aplica os critérios de filtro determinados à coluna fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomItemsFilter(count: number)|Aplica um filtro "Item Inferior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomPercentFilter(percent: number)]|Aplica um filtro "Percentual Inferior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCellColorFilter(color: string)|Aplica um filtro "Cor da Célula" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCustomFilter (criteria1: string, criteria2: string, oper: cadeia de caracteres)|Aplica um filtro "Ícone" à coluna para obter as cadeias de caracteres de critérios fornecidas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyDynamicFilter(criteria: string)|Aplica um filtro "Dinâmico" à coluna.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyFontColorFilter(color: string)|Aplica um filtro "Cor da Fonte" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyIconFilter(icon: Icon)|Aplica um filtro "Ícone" à coluna para obter o ícone fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopItemsFilter(count: number)|Aplica um filtro "Item Superior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopPercentFilter(percent: number)|Aplica um filtro "Percentual Superior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyValuesFilter (valores: ())|Aplica um filtro "Valores" à coluna para obter os valores fornecidos.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > Clear)|Limpa o filtro na coluna determinada.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > color|A cadeia HTML de cor usada para filtrar células. Usada com a filtragem "cellColor" e "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion1|O primeiro critério usado para filtrar os dados. Usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion2|O segundo critério usado para filtrar os dados. Só é usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > dynamicCriteria|Os critérios dinâmicos do conjunto Excel.DynamicFilterCriteria a serem aplicados nessa coluna. Usados com a filtragem "dynamic". Os valores possíveis são: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > filterOn|A propriedade usada pelo filtro para determinar se os valores devem ficar visíveis. Os valores possíveis são: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > operator|O operador usado para combinar o critério 1 e 2 ao usar a filtragem "custom". Os valores possíveis são: "And", "Or".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > values|O conjunto de valores a serem usados como parte da filtragem "values".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relação_ > icon|O ícone usado para filtrar células. Usado com a filtragem "icon".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > date|A data no formato ISO8601 usada para filtrar os dados.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > specificity|Como a data específica deve ser usada para manter os dados. Por exemplo, se a data for 2005-04-02 e a especificidade estiver definida como "mês", a operação de filtragem manterá todas as linhas com uma data do mês de abril de 2009. Os valores possíveis são: Ano, segunda-feira, dia, hora, minuto, segundo.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > formulaHidden|Indica se o Excel ocultará a fórmula para as células no intervalo. Um valor nulo indica que o intervalo inteiro não tem configuração uniforme de fórmula oculta.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > locked|Indica se o Excel bloqueia as células no objeto. Um valor nulo indica que o intervalo inteiro não tem configuração de bloqueio uniforme.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > index|Representa o índice do ícone no conjunto fornecido.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > set|Representa o conjunto do qual ícone faz parte. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > columnHidden|Representa se todas as colunas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > hidden|Representa se todas as células do intervalo atual estão ocultas. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > rowHidden|Representa se todas as linhas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relação_ > sort|Representa a classificação de intervalo do intervalo atual. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > merge(across: bool)|Mescla as células do intervalo em uma região da planilha.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > unmerge()|Desfaz a mesclagem das células do intervalo em células separadas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > columnWidth|Obtém ou define a largura de todas as colunas dentro do intervalo. Se as larguras das colunas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > rowHeight|Obtém ou define a altura de todas as linhas do intervalo. Se as alturas das linhas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relação_ > protection|Retorna o objeto de proteção de formato para um intervalo. Somente leitura.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitColumns()|Altera a largura das colunas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitRows()|Altera a altura das linhas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Propriedade_ > address|Representa as linhas visíveis do intervalo atual.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Método_ > Aplicar (campos: SortField, matchCase: bool, hasHeaders: bool, orientação: cadeia de caracteres, método: cadeia de caracteres)|Execute uma operação de classificação.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > ascending|Indica se a classificação é feita de forma crescente.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > color|Representa a cor que é o destino da condição se a classificação estiver na cor da fonte ou da célula.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > dataOption|Representa as opções de classificação adicionais para esse campo. Os valores possíveis são: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > key|Representa a coluna (ou linha, dependendo da orientação da classificação) em que a condição está. Representado como um deslocamento da primeira coluna (ou linha).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > sortOn|Representa o tipo de classificação dessa condição. Os valores possíveis são: Valor, CellColor, FontColor, Ícone.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relação_ > icon|Representa o ícone que é o destino da condição se a classificação está no ícone da célula.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > sort|Representa a classificação da tabela. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > planilha|A planilha que contém a tabela atual. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > clearFilters()|Limpa todos os filtros aplicados à tabela no momento.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > convertToRange()|Converte a tabela em um intervalo de células normal. Todos os dados são preservados.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > reapplyFilters()|Aplica novamente todos os filtros à tabela.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relação_ > filter|Recupera o filtro aplicado à coluna. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > matchCase|Indica se o uso de maiúsculas ou minúsculas afetou a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > method|Indica o último método de ordenação de caracteres chineses usado para classificar a tabela. Somente leitura. Os valores possíveis são: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relação_ > fields|Representa as condições atuais usadas para a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > Aplicar (campos: SortField, matchCase: bool, método: cadeia de caracteres)|Execute uma operação de classificação.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > Clear)|Limpa a classificação que está na tabela. Essa ação não modifica a ordenação da tabela, mas limpa o estado dos botões do cabeçalho.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > reapply()|Reaplica os parâmetros de classificação atuais à tabela.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > funções|Representa uma instância de aplicativo do Excel que contém essa pasta de trabalho. Somente leitura.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > protection|Retorna o objeto de proteção da planilha para uma planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Propriedade_ > protected|Indica se a planilha está protegida. Somente Leitura. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relação_ > options|Opções de proteção da planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > protect(options: WorksheetProtectionOptions)|Protege uma planilha. Falhará se uma planilha estiver protegida.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > unprotect()|Desprotege uma planilha.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowAutoFilter|Indica a opção de proteção de planilha para permitir a utilização do recurso de filtro automático.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteColumns|Indica a opção de proteção de planilha para permitir a exclusão de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteRows|Indica a opção de proteção de planilha para permitir a exclusão de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatCells|Indica a opção de proteção de planilha para permitir a formatação de células.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatColumns|Indica a opção de proteção de planilha para permitir a formatação de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatRows|Indica a opção de proteção de planilha para permitir a formatação de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertColumns|Indica a opção de proteção de planilha para permitir a inserção de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertHyperlinks|Indica a opção de proteção de planilha para permitir a inserção de hiperlinks.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertRows|Indica a opção de proteção de planilha para permitir a inserção de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowPivotTables|Indica a opção de proteção de planilha para permitir a utilização do recurso de Tabela Dinâmica.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowSort|Indica a opção de proteção de planilha para permitir a utilização do recurso de classificação.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript do Excel 1.1

Excel JavaScript API 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte os tópicos de referência de [API do JavaScript de Excel](/javascript/api/excel) .

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar hosts do Office e requisitos da API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
