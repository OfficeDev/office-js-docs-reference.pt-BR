# <a name="outlook-javascript-api-requirement-sets"></a>Conjuntos de requisito de API do JavaScript do Outlook

Suplementos do Outlook declarar quais versões de API exigem usando o elemento de [requisitos](/javascript/office/manifest/requirements) no seu [manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Os suplementos do Outlook sempre incluem um elemento [Set](/javascript/office/manifest/set) com um atributo `Name` definido como `Mailbox` e um atributo `MinVersion` definido como o conjunto de requisitos mínimo de API que dê suporte aos cenários do suplemento.

Por exemplo, o seguinte trecho do manifesto indica um conjunto de requisitos mínimo de 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas as APIs do Outlook pertencem ao [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) `Mailbox`. O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs que lançamos pertence a uma versão superior. Nem todos os clientes do Outlook serão compatíveis com o conjunto mais recente de APIs, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, será compatível com todas as APIs nesse conjunto.

A especificação de uma versão mínima de conjunto de requisitos controla em quais clientes do Outlook o suplemento aparecerá. Se um cliente não oferece suporte para o conjunto de requisitos mínimos, ele não carrega o suplemento. Por exemplo, se for especificada a versão 1.3 do conjunto de requisitos, significa que o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Usar APIs de conjuntos de requisitos posteriores

A definição de um conjunto de requisito não limitam as APIs disponíveis que pode usar o suplemento. Por exemplo, se o suplemento Especifica requisito definir 1.1, mas ele está sendo executado em um cliente do Outlook que oferece suporte a 1.3, o suplemento pode usar APIs do conjunto de requisito 1.3.

Para usar APIs mais recentes, os desenvolvedores apenas podem verificar sua existência usando técnica de JavaScript padrão:

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Essas verificações não são necessárias para APIs que estão presentes na versão do conjunto de requisitos especificada no manifesto.

## <a name="choosing-a-minimum-requirement-set"></a>Escolher um conjunto de requisitos mínimos

Os desenvolvedores devem usar o conjunto de requisitos mínimos que contém o conjunto essencial de APIs para seu cenário, sem o qual o suplemento não funcionará.

## <a name="clients"></a>Clientes

Os clientes a seguir oferecem suporte para suplementos do Outlook.

| Cliente | Conjuntos de requisitos de API suportados |
| --- | --- |
| Outlook 2019 para Windows | 1.1, 1.2, 1.3, 1,4, 1,5, 1.6 |
| 2019 do Outlook para Mac | 1.1, 1.2, 1.3, 1,4, 1,5, 1.6 |
| Outlook 2016 (Clique para Executar) para Windows | 1.1, 1.2, 1.3, 1,4, 1,5, 1.6, 1.7 |
| Outlook 2016 (MSI) para Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 para Mac | 1.1, 1.2, 1.3, 1,4, 1,5, 1.6 |
| Outlook 2013 para Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook para iPhone | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook para Android | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook na Web (Office 365 e Outlook.com) | 1.1, 1.2, 1.3, 1,4, 1,5, 1.6 |
| Outlook Web App (Exchange 2013 local) | 1.1 |
| Outlook Web App (Exchange 2016 local) | 1.1, 1.2. 1.3 |

> [!NOTE]
> Suporte para 1.3 no Outlook 2013 foi adicionado como parte do [8 de dezembro 2015, atualizar para o Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). Suporte para 1,4 no Outlook 2013 foi adicionado como parte da [13 de setembro, 2016, atualizar para o Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).
