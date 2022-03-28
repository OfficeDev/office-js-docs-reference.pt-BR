---
title: Elemento Hosts no arquivo de manifesto
description: Especifica os aplicativos Office cliente em que o Office o Add-in será ativado.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 060a98229b90d7f6a8c467af21dfef4c63c922d9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485191"
---
# <a name="hosts-element"></a>Elemento Hosts

Especifica os aplicativos Office cliente em que o Office o Add-in será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

## <a name="as-child-of-versionoverrides-element"></a>Como filho do elemento VersionOverrides

As informações nesta seção só se *aplicarão* quando o **elemento Hosts** for filho de [um VersionOverrides](versionoverrides.md).

Esse elemento substitui o **elemento Hosts** no manifesto base.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
