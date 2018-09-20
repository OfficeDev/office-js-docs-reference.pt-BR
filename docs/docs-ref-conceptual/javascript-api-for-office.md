# <a name="javascript-api-for-office"></a>JavaScript API for Office

A API do JavaScript para Office permite criar aplicativos da web que interagem com os modelos de objeto nos aplicativos de host do Office. Seu aplicativo farão referência a biblioteca do Office. js, que é um carregador de script. Biblioteca do Office. js carrega os modelos de objeto que são aplicáveis ao aplicativo do Office que está executando o add-in. Você pode usar os seguintes modelos de objeto JavaScript:

- **APIs comuns** - APIs que foram introduzidos com o **Office 2013**. Isso é carregado para **todos os aplicativos de host do Office** e os conecta o seu aplicativo do suplemento com o aplicativo cliente do Office. O modelo de objeto contém APIs que são específicos para clientes do Office e APIs que são aplicáveis a vários aplicativos de host de cliente do Office. Todo esse conteúdo está sob o **API Shared**. 

  **O Outlook** também usa a sintaxe da API comum. Todos os itens sob o alias Office contém objetos que você pode usar para escrever scripts que interagem com o conteúdo no Office documentos, planilhas, apresentações, itens de email e projetos a partir de seus suplementos do Office. Se seu suplemento será alvo Office 2013 e posterior, você deve usar essas APIs comuns. Esse modelo de objeto usa retornos de chamada.

- **APIs específicas de host** - APIs introduzidas com **Office 2016**. Este modelo de objeto fornece objetos fortemente tipadas específica do host que correspondem aos objetos familiares que você vê ao usar os clientes do Office e representa o futuro do Office JavaScript APIs. As APIs de host específico atualmente incluem a API do JavaScript Word e a API do JavaScript Excel.

## <a name="supported-host-applications"></a>Aplicativos host compatíveis

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API compartilhado](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint e projeto](requirement-sets/powerpoint-and-project-note.md) suporte a suplementos feitos com o API do JavaScript. No entanto, no momento não têm APIs específicas de host. Você interage com esses hosts por meio da API Shared.

Saiba mais sobre [hosts compatíveis e outros requisitos](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Especificações abertas da API

À medida que criamos e desenvolvemos novas APIs para suplementos do Office, nós as disponibilizamos em nossa página [Especificações abertas da API](openspec.md) a fim de obter os seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.

## <a name="see-also"></a>Confira também

- [Referência de API do JavaScript do Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)