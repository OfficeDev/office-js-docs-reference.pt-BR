# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Como a documentação da API JavaScript do Office é gerada

As páginas de documentação de referência do JavaScript para Office são geradas a partir de arquivos de definição de tipo e trechos de exemplo. Esse processo usa uma mistura de ferramentas de código aberto e scripts específicos do repositório. Este documento tem como objetivo tornar os processos deste repositório transparentes para que a comunidade possa se beneficiar melhor e contribuir com esse conteúdo.

## <a name="content-sources"></a>Fontes de conteúdo

Dois tipos de conteúdo são combinados para criar a documentação de referência do Office-JS: definições de tipo e trechos de código. Eles garantem uma cobertura completa da API e dão pequenos exemplos de código em linha.

### <a name="type-definition-files"></a>Arquivos de definição de tipo

Os arquivos de definição de tipo [em Definitivamente Digitados](https://github.com/DefinitelyTyped/DefinitelyTyped) são a única fonte da verdade para a documentação. Qualquer Complemento do Office que usa TypeScript compila usando esses arquivos de definição de tipo. Eles também dão recursos do IntelliSense para desenvolvedores de JavaScript e TypeScript. Ao criar a documentação de referência a partir dessas definições, fornecemos informações mais precisas.

Há quatro arquivos d.ts relevantes que fornecem conteúdo de origem para diferentes subseções dos documentos.

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (As definições de versão.)
  - [Excel (Versão)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (Versão)](https://docs.microsoft.com/javascript/api/word_release)
  - [Subseção OfficeExtensions da API comum](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (As definições de visualização.)
  - [Excel (Visualização)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (Visualização)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (Visualização)](https://docs.microsoft.com/javascript/api/word)
  - [API Comum](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (as definições de tempo de execução de Funções Personalizadas do Excel.)
  - [Funções personalizadas](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (As definições de tempo de execução do office para a plataforma de funções personalizadas.)
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

As versões mais antigas das APIs têm seus próprios arquivos d.ts. Eles são preservados quando um novo conjunto de requisitos de API é lançado. Eles também podem ser gerados usando a ferramenta [Remover Versão.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts) Esses arquivos d.ts antigos são mantidos para que, nas APIs de evento sejam corrigidas ou alteradas, o comportamento original ainda esteja documentado. Isso será útil se você tiver que direcionar uma versão mais antiga da API.

#### <a name="testing-type-definition-file-changes"></a>Teste de alterações de arquivo de definição de tipo

Qualquer alteração de documentação para a API JavaScript do Office é feita editando os quatro arquivos d.ts mencionados acima. No entanto, você pode testar uma alteração antes de enviar um PR para DefinitelyTyped (se precisar, por exemplo, testar como sua formatação será traduzida para markdown) editando o arquivo correspondente em [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) e executando [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). Quando solicitado, selecione a opção "Arquivos locais".

Fazer alterações em um branch remoto desse repo faz com que a plataforma docs.microsoft.com crie uma ramificação de teste. Essa ramificação é renderizada review.docs.microsoft.com, que só pode ser acessada pela equipe interna da Microsoft. Qualquer pessoa que analisar seu PR verificará a precisão do site de revisão.

### <a name="code-snippets"></a>Trechos de código

Trechos de código de exemplo são adicionados às páginas de referência de duas fontes:

- [Amostras do Script Lab](https://github.com/OfficeDev/office-js-snippets)
- [Trechos de Código Local](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Os trechos de código locais estão em arquivos yaml específicos do host. Seu conteúdo é organizado por classe e campo, para que possa ser mapeado para o local apropriado em uma página de referência. A linguagem do trecho de código (JavaScript ou TypeScript) é inferida pelo uso de instruções de espera.

Os trechos de código do Script Lab são retirados de amostras de trabalho. Atualmente, exemplos do Excel, Outlook, PowerPoint e Word são mapeados para seções de documento de referência por meio de [arquivos de mapeamento.](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata) Eles combinam métodos de exemplo individuais com propriedades ou métodos na API. Quando o repositório office-js-snippets é executado, um arquivo yaml que contém todos os `yarn start` trechos mapeados é criado. [](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) Esse arquivo yaml é a entrada na ferramenta de documentação de referência.

## <a name="tooling-pipeline"></a>Pipeline de ferramentas

![Uma imagem mostrando o fluxo de controle de Definitivamente Digitado, para o pré-processador, o Extrator de API, o processador médio, o Documentador de API e até o pós-processador.](ToolingPipeline.png)

Entre as fontes de conteúdo e as páginas finais, o conteúdo da documentação passa por cinco etapas de ferramentas:

1. [Script de pré-processador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Extrator de API](https://api-extractor.com/)
1. [Script Midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Documentador de API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script postprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

O pré-processador pega os arquivos d.ts e os divide em seções específicas do host. Ele executa qualquer limpeza necessária para que as ferramentas subsequentes processem corretamente os dados.

O Extrator de API converte os arquivos d.ts em dados JSON. Isso tokeniza todos os dados do tipo, permitindo uma análise mais fácil.

O processador médio recupera os trechos de código, os emparelha com os hosts apropriados e limpa o crosslink entre objetos do Outlook e da API comum.

O Documentador de API converte os dados JSON em arquivos .yml. Os arquivos .yml são convertidos em markdown pelo Open Publishing System que publica nossos documentos no docs.microsoft.com. O Documentador de API também contém uma extensão específica do Office que insere nossos trechos de código.

O pós-processador limpa o índice e move os arquivos .yml para a [pasta de publicação.](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)

Todas essas cinco etapas são executadas [quando GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) é executado. Esse script também lida com a instalação do módulo de nó, limpa conjuntos de arquivos antigos e versões dos arquivos de Definição de Tipo para cada conjunto de requisitos.
