# <a name="how-the-office-javascript-api-documentation-is-generated"></a><span data-ttu-id="3e110-101">Como a documentação da API JavaScript do Office é gerada</span><span class="sxs-lookup"><span data-stu-id="3e110-101">How the Office JavaScript API documentation is generated</span></span>

<span data-ttu-id="3e110-102">As páginas de documentação de referência JavaScript do Office são geradas de arquivos de definição de tipo e trechos de exemplo.</span><span class="sxs-lookup"><span data-stu-id="3e110-102">The Office JavaScript reference documentation pages are generated from type definition files and example snippets.</span></span> <span data-ttu-id="3e110-103">Esse processo usa uma mistura de ferramentas de código aberto e scripts específicos do repositório.</span><span class="sxs-lookup"><span data-stu-id="3e110-103">This process uses a blend of open source tools and repository-specific scripts.</span></span> <span data-ttu-id="3e110-104">Este documento visa tornar os processos desse repositório transparentes para que a Comunidade possa se beneficiar melhor e contribuir com esse conteúdo.</span><span class="sxs-lookup"><span data-stu-id="3e110-104">This document aims to make the processes of this repository transparent so the community can better benefit from and contribute to this content.</span></span>

## <a name="content-sources"></a><span data-ttu-id="3e110-105">Fontes de conteúdo</span><span class="sxs-lookup"><span data-stu-id="3e110-105">Content sources</span></span>

<span data-ttu-id="3e110-106">Dois tipos de conteúdo são combinados para criar a documentação de referência do Office-JS: definições de tipo e trechos de código.</span><span class="sxs-lookup"><span data-stu-id="3e110-106">Two types of content are combined to create the Office-JS reference documentation: type definitions and code snippets.</span></span> <span data-ttu-id="3e110-107">Isso garante uma cobertura de API completa e fornece exemplos de código em linha pequenas.</span><span class="sxs-lookup"><span data-stu-id="3e110-107">These ensure complete API coverage and give small, inline code samples.</span></span>

### <a name="type-definition-files"></a><span data-ttu-id="3e110-108">Arquivos de definição de tipo</span><span class="sxs-lookup"><span data-stu-id="3e110-108">Type definition files</span></span>

<span data-ttu-id="3e110-109">Os arquivos de definição de tipo em [definitivamente digitados](https://github.com/DefinitelyTyped/DefinitelyTyped) são a única fonte de verdade para a documentação.</span><span class="sxs-lookup"><span data-stu-id="3e110-109">The type definition files on [Definitely Typed](https://github.com/DefinitelyTyped/DefinitelyTyped) are the single source of truth for the documentation.</span></span> <span data-ttu-id="3e110-110">Qualquer suplemento do Office que usa o TypeScript compila usando esses arquivos de definição de tipo.</span><span class="sxs-lookup"><span data-stu-id="3e110-110">Any Office Add-in that uses TypeScript compiles by using these type definition files.</span></span> <span data-ttu-id="3e110-111">Eles também oferecem recursos de IntelliSense para desenvolvedores de JavaScript e TypeScript.</span><span class="sxs-lookup"><span data-stu-id="3e110-111">These also give JavaScript and TypeScript developers IntelliSense capabilities.</span></span> <span data-ttu-id="3e110-112">Ao criar a documentação de referência dessas definições, fornecemos informações mais precisas.</span><span class="sxs-lookup"><span data-stu-id="3e110-112">By building the reference documentation from these definitions, we provide more accurate information.</span></span>

<span data-ttu-id="3e110-113">Há quatro arquivos d. TS relevantes que fornecem conteúdo de origem para subseções diferentes dos documentos.</span><span class="sxs-lookup"><span data-stu-id="3e110-113">There are four relevant d.ts files that provide source content for different subsections of the docs.</span></span>

- <span data-ttu-id="3e110-114">[Office-js/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (As definições de lançamento.)</span><span class="sxs-lookup"><span data-stu-id="3e110-114">[office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (The Release definitions.)</span></span>
  - [<span data-ttu-id="3e110-115">Excel (versão)</span><span class="sxs-lookup"><span data-stu-id="3e110-115">Excel (Release)</span></span>](https://docs.microsoft.com/javascript/api/excel_release)
  - [<span data-ttu-id="3e110-116">OneNote</span><span class="sxs-lookup"><span data-stu-id="3e110-116">OneNote</span></span>](https://docs.microsoft.com/javascript/api/onenote)
  - [<span data-ttu-id="3e110-117">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3e110-117">PowerPoint</span></span>](https://docs.microsoft.com/javascript/api/powerpoint)
  - [<span data-ttu-id="3e110-118">Visio</span><span class="sxs-lookup"><span data-stu-id="3e110-118">Visio</span></span>](https://docs.microsoft.com/javascript/api/visio)
  - [<span data-ttu-id="3e110-119">Word (versão)</span><span class="sxs-lookup"><span data-stu-id="3e110-119">Word (Release)</span></span>](https://docs.microsoft.com/javascript/api/word_release)
  - [<span data-ttu-id="3e110-120">OfficeExtensions subseção de API comum</span><span class="sxs-lookup"><span data-stu-id="3e110-120">OfficeExtensions subsection of the Common API</span></span>](https://docs.microsoft.com/javascript/api/office)
- <span data-ttu-id="3e110-121">[Office-js-Preview/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (As definições de visualização.)</span><span class="sxs-lookup"><span data-stu-id="3e110-121">[office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (The Preview definitions.)</span></span>
  - [<span data-ttu-id="3e110-122">Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="3e110-122">Excel (Preview)</span></span>](https://docs.microsoft.com/javascript/api/excel)
  - [<span data-ttu-id="3e110-123">Outlook (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="3e110-123">Outlook (Preview)</span></span>](https://docs.microsoft.com/javascript/api/outlook)
  - [<span data-ttu-id="3e110-124">Word (visualização)</span><span class="sxs-lookup"><span data-stu-id="3e110-124">Word (Preview)</span></span>](https://docs.microsoft.com/javascript/api/word)
  - [<span data-ttu-id="3e110-125">API Comum</span><span class="sxs-lookup"><span data-stu-id="3e110-125">Common API</span></span>](https://docs.microsoft.com/javascript/api/office)
- <span data-ttu-id="3e110-126">[Custom-Functions-Runtime](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (As definições de tempo de execução de funções personalizadas do Excel.)</span><span class="sxs-lookup"><span data-stu-id="3e110-126">[custom-functions-runtime](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts)(The Excel Custom Functions runtime definitions.)</span></span>
  - [<span data-ttu-id="3e110-127">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="3e110-127">Custom Functions</span></span>](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- <span data-ttu-id="3e110-128">[Office – tempo de execução](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (As definições de tempo de execução do Office para a plataforma de funções personalizadas.)</span><span class="sxs-lookup"><span data-stu-id="3e110-128">[office-runtime](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts)(The office runtime definitions for the Custom Functions platform.)</span></span>
  - [<span data-ttu-id="3e110-129">Tempo de execução do Office</span><span class="sxs-lookup"><span data-stu-id="3e110-129">Office Runtime</span></span>](https://docs.microsoft.com/javascript/api/office-runtime)

<span data-ttu-id="3e110-130">Versões mais antigas das APIs têm seus próprios arquivos d. TS.</span><span class="sxs-lookup"><span data-stu-id="3e110-130">Older versions of the APIs have their own d.ts files.</span></span> <span data-ttu-id="3e110-131">Eles são preservados quando um novo conjunto de requisitos de API é lançado.</span><span class="sxs-lookup"><span data-stu-id="3e110-131">These are preserved when a new API requirement set is released.</span></span> <span data-ttu-id="3e110-132">Elas também podem ser geradas usando a [ferramenta removedor de versão](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts).</span><span class="sxs-lookup"><span data-stu-id="3e110-132">They can also be generated using the [Version Remover tool](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts).</span></span> <span data-ttu-id="3e110-133">Esses arquivos d. TS antigos são mantidos de forma que as APIs de eventos sejam corrigidas ou alteradas, o comportamento original ainda está documentado.</span><span class="sxs-lookup"><span data-stu-id="3e110-133">These old d.ts files are maintained so that in the event APIs are patched or altered, the original behavior is still documented.</span></span> <span data-ttu-id="3e110-134">Isso será útil se você tiver que direcionar uma versão mais antiga da API.</span><span class="sxs-lookup"><span data-stu-id="3e110-134">This is useful if you have to target an older version of the API.</span></span>

### <a name="code-snippets"></a><span data-ttu-id="3e110-135">Trechos de código</span><span class="sxs-lookup"><span data-stu-id="3e110-135">Code snippets</span></span>

<span data-ttu-id="3e110-136">Trechos de código de exemplo são adicionados às páginas de referência de duas fontes:</span><span class="sxs-lookup"><span data-stu-id="3e110-136">Code example snippets are added to the reference pages from two sources:</span></span>

- [<span data-ttu-id="3e110-137">Amostras de laboratório de script</span><span class="sxs-lookup"><span data-stu-id="3e110-137">Script Lab Samples</span></span>](https://github.com/OfficeDev/office-js-snippets)
- [<span data-ttu-id="3e110-138">Trechos de código locais</span><span class="sxs-lookup"><span data-stu-id="3e110-138">Local Code Snippets</span></span>](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

<span data-ttu-id="3e110-139">Os trechos de código locais estão em arquivos YAML específicos do host.</span><span class="sxs-lookup"><span data-stu-id="3e110-139">The local snippets are in host-specific yaml files.</span></span> <span data-ttu-id="3e110-140">O conteúdo é organizado por classe e campo, para que possa ser mapeado para o local apropriado em uma página de referência.</span><span class="sxs-lookup"><span data-stu-id="3e110-140">Their content is organized by class and field, so it can be mapped to the appropriate place in a reference page.</span></span> <span data-ttu-id="3e110-141">O idioma do trecho de código (JavaScript ou TypeScript) é inferido pelo uso de instruções Await.</span><span class="sxs-lookup"><span data-stu-id="3e110-141">The language of the snippet (JavaScript or TypeScript) is inferred by the use of await statements.</span></span>

<span data-ttu-id="3e110-142">Os trechos de laboratório de script são retirados de exemplos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="3e110-142">The Script Lab snippets are pulled from working samples.</span></span> <span data-ttu-id="3e110-143">Atualmente, as amostras do Excel e do Word são mapeadas para seções de documento de referência por meio [de um par de arquivos de mapeamento](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata).</span><span class="sxs-lookup"><span data-stu-id="3e110-143">Currently, Excel and Word samples are mapped to reference document sections through a [pair of mapping files](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata).</span></span> <span data-ttu-id="3e110-144">Eles correspondem a métodos de amostra individuais para propriedades ou métodos na API.</span><span class="sxs-lookup"><span data-stu-id="3e110-144">These match individual sample methods to properties or methods in the API.</span></span> <span data-ttu-id="3e110-145">Quando o repositório Office-js-Snippets é `yarn start` executado, [um arquivo YAML](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) contendo todos os trechos de código mapeados é criado.</span><span class="sxs-lookup"><span data-stu-id="3e110-145">When the office-js-snippets repository's `yarn start` runs, [a yaml file](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) containing all the mapped snippets is created.</span></span> <span data-ttu-id="3e110-146">Este arquivo do YAML é a entrada para a ferramenta de documentação de referência.</span><span class="sxs-lookup"><span data-stu-id="3e110-146">This yaml file is the input into the reference documentation tooling.</span></span>

## <a name="tooling-pipeline"></a><span data-ttu-id="3e110-147">Pipeline de ferramentas</span><span class="sxs-lookup"><span data-stu-id="3e110-147">Tooling pipeline</span></span>

![Uma imagem que mostra o fluxo de controle de definitivamente digitado, para o pré-processador, o Extrador de API, o Documentador de API e o perprocessador.](ToolingPipeline.png)

<span data-ttu-id="3e110-149">Entre as fontes de conteúdo e as páginas finais, o conteúdo da documentação passa por quatro etapas de ferramenta:</span><span class="sxs-lookup"><span data-stu-id="3e110-149">Between the content sources and the final pages, the documentation content goes through four tooling steps:</span></span>

1. [<span data-ttu-id="3e110-150">Script de pré-processador</span><span class="sxs-lookup"><span data-stu-id="3e110-150">Preprocessor script</span></span>](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [<span data-ttu-id="3e110-151">Extrator de API</span><span class="sxs-lookup"><span data-stu-id="3e110-151">API Extractor</span></span>](https://api-extractor.com/)
1. [<span data-ttu-id="3e110-152">Documentador de API</span><span class="sxs-lookup"><span data-stu-id="3e110-152">API Documenter</span></span>](https://github.com/microsoft/web-build-tools/blob/master/apps/api-documenter/README.md)
1. [<span data-ttu-id="3e110-153">Script do co-processador</span><span class="sxs-lookup"><span data-stu-id="3e110-153">Postprocessor script</span></span>](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

<span data-ttu-id="3e110-154">O pré-processador usa os arquivos d. TS e os divide em seções específicas do host.</span><span class="sxs-lookup"><span data-stu-id="3e110-154">The preprocessor takes the d.ts files and splits them into host-specific sections.</span></span> <span data-ttu-id="3e110-155">Ele realiza qualquer limpeza necessária para as ferramentas subsequentes para processar corretamente os dados.</span><span class="sxs-lookup"><span data-stu-id="3e110-155">It performs any cleanup necessary for the subsequent tools to properly process the data.</span></span>

<span data-ttu-id="3e110-156">O extrator de API converte os arquivos d. TS em dados JSON.</span><span class="sxs-lookup"><span data-stu-id="3e110-156">API Extractor converts the d.ts files into JSON data.</span></span> <span data-ttu-id="3e110-157">Isso tokenizes todos os dados de tipo, permitindo uma análise mais fácil.</span><span class="sxs-lookup"><span data-stu-id="3e110-157">This tokenizes all the type data, allowing for easier parsing.</span></span>

<span data-ttu-id="3e110-158">O Documentador de API converte os dados JSON em arquivos. yml.</span><span class="sxs-lookup"><span data-stu-id="3e110-158">API Documenter converts the JSON data into .yml files.</span></span> <span data-ttu-id="3e110-159">Os arquivos. yml são convertidos para redução pelo sistema de publicação aberto que publica nossos documentos para o docs.microsoft.com.</span><span class="sxs-lookup"><span data-stu-id="3e110-159">The .yml files are converted to markdown by the Open Publishing System that publishes our docs to docs.microsoft.com.</span></span> <span data-ttu-id="3e110-160">O Documentador de API também contém uma extensão específica do Office que insere nossos trechos de código.</span><span class="sxs-lookup"><span data-stu-id="3e110-160">API Documenter also contains an Office-specific extension that inserts our code snippets.</span></span>

<span data-ttu-id="3e110-161">O co-processador limpa o Sumário e move os arquivos. yml para a [pasta de publicação](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).</span><span class="sxs-lookup"><span data-stu-id="3e110-161">The postprocessor cleans up the table of contents and moves the .yml files into the [publishing folder](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).</span></span>

<span data-ttu-id="3e110-162">Todas as quatro etapas são executadas quando o [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) é executado.</span><span class="sxs-lookup"><span data-stu-id="3e110-162">All four of these steps are performed when [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) is run.</span></span> <span data-ttu-id="3e110-163">Esse script também trata da instalação do módulo de nó e limpa os conjuntos de arquivos antigos.</span><span class="sxs-lookup"><span data-stu-id="3e110-163">That script also handles node module installation and cleans out old file sets.</span></span>