# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Como a documentação da API JavaScript do Office é gerada

As páginas de documentação de referência JavaScript do Office são geradas de arquivos de definição de tipo e trechos de exemplo. Esse processo usa uma mistura de ferramentas de código aberto e scripts específicos do repositório. Este documento visa tornar os processos desse repositório transparentes para que a Comunidade possa se beneficiar melhor e contribuir com esse conteúdo.

## <a name="content-sources"></a>Fontes de conteúdo

Dois tipos de conteúdo são combinados para criar a documentação de referência do Office-JS: definições de tipo e trechos de código. Isso garante uma cobertura de API completa e fornece exemplos de código em linha pequenas.

### <a name="type-definition-files"></a>Arquivos de definição de tipo

Os arquivos de definição de tipo em [definitivamente digitados](https://github.com/DefinitelyTyped/DefinitelyTyped) são a única fonte de verdade para a documentação. Qualquer suplemento do Office que usa o TypeScript compila usando esses arquivos de definição de tipo. Eles também oferecem recursos de IntelliSense para desenvolvedores de JavaScript e TypeScript. Ao criar a documentação de referência dessas definições, fornecemos informações mais precisas.

Há quatro arquivos d. TS relevantes que fornecem conteúdo de origem para subseções diferentes dos documentos.

- [Office-js/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (as definições de versão.)
  - [Excel (versão)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (versão)](https://docs.microsoft.com/javascript/api/word_release)
  - [OfficeExtensions subseção de API comum](https://docs.microsoft.com/javascript/api/office)
- [Office-js-Preview/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (as definições de visualização.)
  - [Excel (versão prévia)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (versão prévia)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (visualização)](https://docs.microsoft.com/javascript/api/word)
  - [API Comum](https://docs.microsoft.com/javascript/api/office)
- [Custom-Functions-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (as definições de tempo de execução de funções personalizadas do Excel.)
  - [Funções personalizadas](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [Office-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (as definições de tempo de execução do Office para a plataforma de funções personalizadas.)
  - [Tempo de execução do Office](https://docs.microsoft.com/javascript/api/office-runtime)

Versões mais antigas das APIs têm seus próprios arquivos d. TS. Eles são preservados quando um novo conjunto de requisitos de API é lançado. Elas também podem ser geradas usando a [ferramenta removedor de versão](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts). Esses arquivos d. TS antigos são mantidos de forma que as APIs de eventos sejam corrigidas ou alteradas, o comportamento original ainda está documentado. Isso será útil se você tiver que direcionar uma versão mais antiga da API.

#### <a name="testing-type-definition-file-changes"></a>Alterações de arquivo de definição de tipo de teste

Qualquer alteração de documentação para a API JavaScript do Office é feita com a edição dos quatro arquivos d. TS mencionados acima. No entanto, você pode testar uma alteração antes de enviar um PR para DefinitelyTyped (se precisar, por exemplo, testar como sua formatação será traduzida para redução) editando o arquivo correspondente em [Generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) e executando o [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). Quando solicitado, selecione a opção "arquivos locais".

Enviar alterações a uma filial remota desse repositório faz com que a plataforma docs.microsoft.com Construa uma ramificação de teste. Essa ramificação é renderizada no review.docs.microsoft.com, que só é acessível pela equipe interna da Microsoft. Qualquer pessoa que estiver examinando o seu PR verificará a precisão do site de análise.

### <a name="code-snippets"></a>Trechos de código

Trechos de código de exemplo são adicionados às páginas de referência de duas fontes:

- [Amostras de laboratório de script](https://github.com/OfficeDev/office-js-snippets)
- [Trechos de código locais](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Os trechos de código locais estão em arquivos YAML específicos do host. O conteúdo é organizado por classe e campo, para que possa ser mapeado para o local apropriado em uma página de referência. O idioma do trecho de código (JavaScript ou TypeScript) é inferido pelo uso de instruções Await.

Os trechos de laboratório de script são retirados de exemplos de trabalho. Atualmente, as amostras do Excel, Outlook e do Word são mapeadas para consultar seções de documento através de [arquivos de mapeamento](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata). Eles correspondem a métodos de amostra individuais para propriedades ou métodos na API. Quando o repositório Office-js-Snippets é `yarn start` executado, [um arquivo YAML](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) contendo todos os trechos de código mapeados é criado. Este arquivo do YAML é a entrada para a ferramenta de documentação de referência.

## <a name="tooling-pipeline"></a>Pipeline de ferramentas

![Uma imagem mostrando o fluxo de controle de definitivamente digitado, para o pré-processador, extrator de API, midprocessor, documentador de API e até o co-processador.](ToolingPipeline.png)

Entre as fontes de conteúdo e as páginas finais, o conteúdo da documentação passa por cinco etapas de ferramenta:

1. [Script de pré-processador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Extrator de API](https://api-extractor.com/)
1. [Script Midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Documentador de API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script do co-processador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

O pré-processador usa os arquivos d. TS e os divide em seções específicas do host. Ele realiza qualquer limpeza necessária para as ferramentas subsequentes para processar corretamente os dados.

O extrator de API converte os arquivos d. TS em dados JSON. Isso tokenizes todos os dados de tipo, permitindo uma análise mais fácil.

O midprocessor recupera os trechos de código e os emparelha com os hosts adequados e limpa o crosslinking entre os objetos do Outlook e da API comum.

O Documentador de API converte os dados JSON em arquivos. yml. Os arquivos. yml são convertidos para redução pelo sistema de publicação aberto que publica nossos documentos para o docs.microsoft.com. O Documentador de API também contém uma extensão específica do Office que insere nossos trechos de código.

O co-processador limpa o Sumário e move os arquivos. yml para a [pasta de publicação](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).

Todas as cinco etapas são executadas quando o [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) é executado. Esse script também trata da instalação do módulo de nó, limpa os conjuntos de arquivos antigos e os arquivos de definição de tipo de versão para cada conjunto de requisitos.
