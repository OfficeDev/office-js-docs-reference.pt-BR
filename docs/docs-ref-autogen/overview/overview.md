---
layout: LandingPage
ms.topic: landing-page
title: Referência da API JavaScript do Office
description: APIs JavaScript para Office por host e versão.
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 10/14/2020
ms.openlocfilehash: 04a686fe17222e6ddff53bd5e90f96d2e765da70
ms.sourcegitcommit: b292535480a6b95e10ebbdd7da30f2858d8b7b82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626327"
---
# <a name="office-add-ins-javascript-api-reference"></a>Referência da API JavaScript para suplementos do Office

A API JavaScript para Office permite que você crie aplicativos Web que interajam com os modelos de objeto em aplicativos host do Office. Use esta seção para saber mais sobre as classes, os métodos e outros tipos disponíveis para a criação de Suplementos do Office.

Veja a seguir uma lista de APIs com os [Aplicativos de host do Office som suporte](/office/dev/add-ins/overview/office-add-in-availability). Os links comuns da API incluem todas as APIs não atribuídas a um host em particular (conforme detalhado no [Conjunto de requisitos da API Comum do Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)). Outros itens vinculam a uma versão da documentação de referência da API para esse host com base em um conjunto de requisitos. A documentação de referência será versionada para incluir todas as APIs até a versão do conjunto de requisitos (por exemplo, o ExcelApi 1.3 mostrará APIs do ExcelApi 1.1, 1.2, 1.3, bem como as APIs comuns).

`ExcelApiOnline 1.1` é um conjunto especial de requisitos. Ele contém as APIs mais recentes do Excel na Web, no entanto, essas APIs podem ainda não ser suportadas em todas as plataformas. Confira o [Conjunto de requisitos apenas online da API JavaScript do Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) para obter mais informações.

> [!TIP]
> Você pode alterar a versão de uma página de referência usando o menu suspenso de seleção de filtro acima do sumário a qualquer momento. Se a página não existir nessa versão específica, você retornará à versão atual.

<h2>Hosts do Office</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do Excel</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">Visualização do ExcelApi</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.12">ExcelApi 1.12</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do Outlook</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Visualização de caixa de correio</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.9">Caixa de correio 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Caixa de correio 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Caixa de correio 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Caixa de correio 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Caixa de correio 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Caixa de correio 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Caixa de correio 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Caixa de correio 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Caixa de correio 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do Word</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">Visualização do WordApi</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do OneNote</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do Visio</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do PowerPoint</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-1.1">APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs do Project</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">Somente APIs comuns</a></li>
                </ul>
            </div>
        </a>
    </li>
</ul>

> [!NOTE]
> Se você estiver procurando APIs JavaScript para desenvolver Scripts do Office, acesse a [Referência de API de Scripts do Office](/javascript/api/office-scripts/overview).

## <a name="see-also"></a>Confira também

- [Sobre os Suplementos do Office](/office/dev/add-ins/overview)
- [Disponibilidade de host e plataforma para suplementos do Office](/office/dev/add-ins/overview/office-add-in-availability)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Explore a API JavaScript do Office usando o Script Lab](/office/dev/add-ins/overview/explore-with-script-lab).
