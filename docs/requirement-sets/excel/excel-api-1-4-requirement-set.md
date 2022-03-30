---
title: Excel conjunto de requisitos da API JavaScript 1.4
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 555e75fe4d7a210e22a74081b29869a2dc17b0cb
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503467"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

## <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` - Itens com escopo de planilha ou pasta de trabalho.
* `worksheet` - Retorna a planilha na qual o item nomeado tem escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)` - Adiciona um novo nome à coleção do escopo determinado.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Adiciona um novo nome à coleção do escopo determinado usando a localidade do usuário para a fórmula.

## <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs incluem `getItem()` obter a entrada de configuração por `add()` meio da chave e adicionar o par de configuração key:value especificado à lista de trabalho.

## <a name="others"></a>Outros

* De definir o nome da coluna da tabela.
* Adicione uma coluna de tabela ao final da tabela.
* Adicione várias linhas a uma tabela de cada vez.
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Os [\*métodos e propriedades OrNullObject](/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties): Essa funcionalidade permite obter um objeto usando uma chave. Se o objeto não existir, a propriedade do `isNullObject` objeto retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto sem precisar lidar com ele por meio do tratamento de exceção. Um `*OrNullObject` método está disponível na maioria dos objetos da coleção.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.4. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.4 ou anterior, consulte Excel APIs no conjunto de requisitos [1.4 ou anterior](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_4.md)]

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
