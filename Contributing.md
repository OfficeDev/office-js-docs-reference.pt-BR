# <a name="contribute-to-this-documentation"></a>Contribuir para esta documentação

Agradecemos seu interesse em nossa documentação! As melhores maneiras de ajudar a melhorar essa documentação para todos são:

* [Relatar problemas ou solicitar exemplos adicionais.](#report-issues)
* [Adicione ou edite novos exemplos.](#add-or-edit-samples)

## <a name="report-issues"></a>Relatar problemas

A maior parte do conteúdo desse repo é gerada a partir de [arquivos de definição de tipo](https://github.com/DefinitelyTyped/DefinitelyTyped). Como tal, as correções aqui provavelmente serão substituídas. Informe problemas ou faça solicitações de documentação adicional [por meio GitHub Problemas](https://github.com/OfficeDev/office-js-docs-reference/issues). Para fazer isso, vá para a seção **Comentários** na parte inferior do artigo afetado e selecione **Esta** página para criar um GitHub problema. Como alternativa, crie um novo problema diretamente [no GitHub](https://github.com/OfficeDev/office-js-docs-reference/issues/new).

## <a name="add-or-edit-new-samples"></a>Adicionar ou editar novos exemplos

> **Importante**: o conteúdo de referência neste repositório é gerado automaticamente. Leia as [anotações de ferramentas de documentação](https://github.com/OfficeDev/office-js-docs-reference/blob/main/DocumentationToolingNotes.md) se você estiver planejando editar arquivos não mencionados nesta seção.

Exemplos de código encontrados na documentação deste repo vêm de dois locais:

* [**/docs/code-snippets**](https://github.com/OfficeDev/office-js-docs-reference/tree/main/docs/code-snippets)
* [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets) (A origem de Script Lab amostras)

As contribuições para o Script Lab de exemplo são abordadas [no readme desse repo](https://github.com/OfficeDev/office-js-snippets#readme).

### <a name="how-to-add-or-edit-sample-code"></a>Como adicionar ou editar código de exemplo

Os [**arquivos /docs/code-snippets/\*-snippets.yaml**](https://github.com/OfficeDev/office-js-docs-reference/tree/main/docs/code-snippets) neste repo contêm amostras para cada aplicativo e as APIs Comuns. Cada tecla YAML em um desses arquivos YAML é a UID de um campo na documentação. As teclas seguem o formato `<Application>.<ClassName>#<fieldName>:member<overloadNumber(methods only)>`. Adicione seu trecho de código de exemplo como o valor sob a chave do campo que você está usando em um exemplo. Use os seguintes mecanismos git para enviar alterações ao repositório.

* Para fazer pequenas alterações em um exemplo, [contribua usando GitHub](#contribute-using-github).
* Para fazer grandes alterações, [contribua usando o Git](#contribute-using-git).

### <a name="tips-on-making-good-samples"></a>Dicas em fazer boas amostras

Revisaremos cada RP para garantir que atenda aos padrões da Microsoft. Para ajudar a garantir que seu exemplo seja integrado à documentação de referência, lembre-se das dicas a seguir.

* Limite o escopo. Código de exemplo que se concentra em um pequeno cenário ou uso de API ajuda os leitores a entender como essa API é usada. (Amostras maiores devem ir para [Script Lab](https://github.com/OfficeDev/office-js-snippets) ou o [repo de exemplos de código](https://github.com/officedev/office-add-in-samples).)
* Inclua um breve comentário explicando o comportamento pretendido.
* Estado claro da entrada ou estado esperado.

### <a name="contribute-using-github"></a>Contribuir usando o GitHub

Use GitHub para contribuir com essa documentação sem precisar clonar o repo para sua área de trabalho. Essa é a maneira mais fácil de criar uma solicitação pull neste repositório. Use este método para fazer uma pequena alteração que não envolva alterações de código.

O uso desse método permite que você contribua para um artigo por vez.

#### <a name="to-contribute-using-github"></a>Para contribuir usando GitHub

1. Encontre o artigo para o GitHub.
1. Depois de entrar no artigo no GitHub, entre no GitHub (para obter uma conta [gratuita, Participe GitHub](https://github.com/join)).
1. Escolha o **ícone de lápis** (edite o arquivo em seu bifurcação deste projeto) e faça suas alterações na janela editar<>**arquivo Edit** .
1. Role até a parte inferior e insira uma descrição.
1. Escolha **Propor alteração de arquivoCriar** >  **solicitação de pull**.

Agora você enviou com êxito uma solicitação de pull. Normalmente, as solicitações de pull são revisadas dentro de 10 dias úteis.

### <a name="contribute-using-git"></a>Contribuir usando o Git

Use o Git para contribuir com alterações substantivas, como:

* Código de contribuição.
* Contribuir com alterações que afetam o significado.
* Contribuindo com grandes alterações no texto.
* Adicionando novos tópicos.

#### <a name="to-contribute-using-git"></a>Para contribuir usando o Git

1. Se você não tiver uma conta GitHub, desarmá-la no [GitHub](https://github.com/join).
1. Depois de ter uma conta, instale o Git no computador. Siga as etapas no tutorial [Configurar Git] .
1. Para enviar uma solicitação pull usando o Git, siga as etapas em [Usar GitHub, Git e este repositório](#use-github-git-and-this-repository).
1. Você será solicitado a assinar o Contrato de Licença do Colaborador se estiver:

    * Um membro do grupo Microsoft Open Technologies.
    * Um colaborador que não trabalha para a Microsoft.

Como membro da comunidade, você deve assinar o Contrato de Licença de Contribuição (CLA) antes de poder contribuir com grandes envios para um projeto. Você só precisa concluir e enviar a documentação uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

Assinar o CLA não concede direitos de confirmação ao repositório principal, mas significa que as equipes de Publicação de Conteúdo de Desenvolvedor Office e Office desenvolvedores poderão revisar e aprovar suas contribuições. Você é creditado por seus envios.

## <a name="use-github-git-and-this-repository"></a>Use o GitHub, o Git e este repositório

**Observação:** a maior parte das informações desta seção pode ser encontrada nos artigos de [ajuda no GitHub]. Se você estiver familiarizado com Git e GitHub, pule para a seção Contribuir e [editar](#contribute-and-edit-content) conteúdo para saber mais sobre as especificações do fluxo de código/conteúdo deste repositório.

### <a name="to-set-up-your-fork-of-the-repository"></a>Para configurar seu bifurcação do repositório

1. Configure uma conta GitHub para que você pode contribuir para esse projeto. Se você ainda não fez isso, [vá para GitHub](https://github.com/join) e faça isso agora.
1. Instale o Git em seu computador. Siga as etapas no tutorial [Configurar Git] .
1. Crie sua própria bifurcação para este repositório. Para fazer isso, na parte superior da página, escolha o **botão Bifurcação** .
1. Copie sua bifurcação para o computador. Para fazer isso, abra Git Bash. No prompt de comando, digite:

        git clone https://github.com/<your user name>/office-js-docs-reference.git

    Em seguida, crie uma referência para o repositório raiz inserindo esses comandos:

        cd office-js-docs-reference
        git remote add upstream https://github.com/OfficeDev/office-js-docs-reference.git
        git fetch upstream

Parabéns! Agora seu repositório está configurado. Você não precisará repetir essas etapas novamente.

### <a name="contribute-and-edit-content"></a>Contribuir e editar o conteúdo

Para tornar o processo de contribuição o mais contínuo possível, siga estas etapas.

#### <a name="to-contribute-and-edit-content"></a>Para contribuir e editar conteúdo

1. Crie uma nova ramificação.
1. Adicione novo conteúdo ou edite o conteúdo existente.
1. Envie uma solicitação pull para o repositório principal.
1. Exclua a ramificação.

**Importante**: limite cada filial a um único conceito/artigo para simplificar o fluxo de trabalho e reduzir a chance de conflitos de mesclagem. O conteúdo apropriado para um novo branch inclui:

* Um novo artigo.
* Edições ortográficas e gramaticais.
* Aplicar uma única alteração de formatação em um grande conjunto de artigos (por exemplo, aplicar um novo rodapé de direitos autorais).

#### <a name="to-create-a-new-branch"></a>Para criar um novo branch

1. Abra Git Bash.
1. No prompt de comando Git Bash, digite `git pull upstream main:<new branch name>`. Isso cria uma nova filial localmente copiada da filial principal mais recente do OfficeDev.
1. No prompt de comando Git Bash, digite `git push origin <new branch name>`. Isso alerta GitHub para a nova filial. Agora você deverá surgir a nova ramificação na sua bifurcação do repositório no GitHub.
1. No prompt de comando Git Bash, digite para `git checkout <new branch name>` alternar para sua nova filial.

#### <a name="add-new-content-or-edit-existing-content"></a>Adicionar novo conteúdo ou editar o conteúdo existente

Navegue até o repositório em seu computador usando o Explorador de Arquivos. Os arquivos do repositório estão em `C:\Users\<yourusername>\office-js-docs-reference`.

Para editar arquivos, abra-os em um editor de sua escolha e modifique-os. Para criar um novo arquivo, use o editor de sua escolha e salve o novo arquivo no local apropriado em sua cópia local do repositório. Durante o trabalho, salve seu trabalho com frequência.

Os arquivos em `C:\Users\<yourusername>\office-js-docs-reference` são uma cópia de trabalho da nova ramificação que você criou no repositório local. Qualquer que seja a alteração você faça nessa pasta, ela só afetará o repositório local quando você confirmar uma alteração. Para confirmação de uma alteração no repositório local, digite os seguintes comandos no GitBash.

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

O comando `add` adiciona suas alterações para uma área de preparo em preparação para confirmá-las no repositório. O período após o `add` comando especifica que você deseja estágio de todos os arquivos que você adicionou ou modificou, verificando subpastas recursivamente. (Caso você não queira confirmar todas as alterações, é possível adicionar arquivos específicos. Você também pode desfazer uma confirmação. Para obter ajuda, digite `git add -help` ou `git status`.)

O comando `commit` aplica as alterações preparadas ao repositório. A opção `-m` significa que você está fornecendo o comentário de confirmação na linha de comando. As opções -v e -a podem ser omitidas. A opção -v é para saída detalhada do comando e -a faz o que você já fez com o comando add.

Você pode se comprometer várias vezes enquanto estiver fazendo seu trabalho ou pode se comprometer uma vez quando terminar.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar uma solicitação pull para o repositório principal.

Quando terminar o trabalho e estiver pronto para mesclar no repositório principal, siga estas etapas.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar uma solicitação pull ao repositório principal

1. No prompt de comando Git Bash, digite `git push origin <new branch name>`. Em seu repositório local, `origin` refere-se ao repositório do GitHub a partir do qual você clonou o repositório local. Esse comando coloca o estado atual do sua nova ramificação, incluindo todas as confirmações feitas nas etapas anteriores, na ramificação do GitHub.
1. No site do GitHub, navegue em sua bifurcação para a nova ramificação.
1. Escolha o **botão Puxar Solicitação** na parte superior da página.
1. Verifique se o branch Base é `OfficeDev/office-js-docs-reference@main` e o branch Head é `<your username>/office-js-docs-reference@<branch name>`.
1. Escolha o **botão Atualizar Intervalo de Confirmação** .
1. Adicione um título à sua solicitação de pull e descreva todas as alterações que você está fazendo.
1. Envie a solicitação pull.

Um dos administradores do site processará sua solicitação de pull. Sua solicitação pull será a tona no site de referência do [OfficeDev/office-js-docs em](https://github.com/OfficeDev/office-js-docs-reference/pulls) **Solicitações pull**. Quando a solicitação pull for aceita, o problema será resolvido.

#### <a name="create-a-new-branch-after-merge"></a>Criar uma nova ramificação após a mesclagem

Depois que uma ramificação é mesclada com êxito (ou seja, sua solicitação pull é aceita), não continue trabalhando nesse branch local. Isso pode levar a conflitos de mesclagem se você enviar outra solicitação pull. Para fazer outra atualização, crie uma nova ramificação local a partir da ramificação upstream mesclada com êxito e exclua sua ramificação local inicial.

Por exemplo, se sua filial local X foi mesclada com êxito no branch principal do OfficeDev/office-js-docs-reference e você deseja fazer atualizações adicionais para o conteúdo mesclado. Crie uma nova filial local, X2, a partir do branch principal officeDev/office-js-docs-reference. Para fazer isso, abra o GitBash e execute os seguintes comandos.

    cd office-js-docs-reference
    git pull upstream main:X2
    git push origin X2

Agora você tem cópias locais (em uma nova filial local) do trabalho enviado na ramificação X. A ramificação X2 também contém todos os trabalhos que outros autores mesclaram, portanto, se seu trabalho depender do trabalho de outras pessoas (por exemplo, imagens compartilhadas), ele estará disponível no novo branch. Você pode verificar se o trabalho anterior (e o trabalho de outras pessoas) está no branch fazendo check-out do novo branch...

    git checkout X2

... e verificando o conteúdo. (O `checkout` comando atualiza os arquivos `C:\Users\<yourusername>\office-js-docs-reference` no estado atual da ramificação X2.) Depois de fazer check-out da nova filial, você pode fazer atualizações para o conteúdo e confirma-los como de costume. No entanto, para evitar trabalhar na ramificação mesclada (X) por engano, o melhor a fazer será excluí-la (confira a seguinte seção: **Excluir uma ramificação**).

#### <a name="delete-a-branch"></a>Excluir uma ramificação

Depois que suas alterações são mescladas com êxito no repositório principal, exclua a ramificação que você usou porque não precisa mais dele.  Qualquer trabalho adicional deve ser feito em um novo branch.  

#### <a name="to-delete-a-branch"></a>Para excluir um branch

1. No prompt de comando Git Bash, digite `git checkout main`. Isso garante que você não fique na ramificação a ser excluída (o que não é permitido).
1. Em seguida, no prompt de comando, digite `git branch -d <branch name>`. Isso exclui a ramificação em seu computador somente se tiver sido mesclada com êxito ao repositório upstream. (Você pode substituir esse comportamento com o sinalizador `–D`, mas primeiro certifique-se de que deseja fazer isso).
1. Por fim, digite `git push origin :<branch name>` no comando prompt (um espaço antes dos dois pontos e nenhum espaço depois deles).   Essa ação excluirá a ramificação em uma bifurcação do github.  

Parabéns, você contribuiu com êxito para o projeto!

## <a name="faq"></a>Perguntas frequentes

### <a name="how-do-i-get-a-github-account"></a>Como obter uma conta do GitHub?

Preencha o formulário em [Ingressar no GitHub](https://github.com/join) para abrir uma conta gratuita do GitHub.

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Onde posso obter um Contrato de Licença do Colaborador?

Um aviso será automaticamente enviado para você informando que é preciso assinar o CLA (Contrato de Licença do Colaborador) se sua solicitação de recebimento exigir um.

Como membro da comunidade, **você deve assinar o CLA (Contrato de Licença do Colaborador) antes de poder contribuir com envios volumosos para esse projeto**. Você só precisa concluir e enviar a documentação uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

### <a name="what-happens-with-my-contributions"></a>O que acontece com minhas contribuições?

Quando você enviar suas alterações, por meio de uma solicitação pull, nossa equipe será notificada e revisará sua solicitação de pull. Você receberá notificações sobre sua solicitação de GitHub; você também poderá ser notificado por alguém da nossa equipe se precisarmos de mais informações. Se sua solicitação de pull for aprovada, atualizaremos a documentação. Reservamo-nos o direito de editar seu envio para questões legais, de estilo, de clareza ou de outros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Posso me tornar um aprovador para as solicitações de GitHub pull deste repositório?

Atualmente, não estamos permitindo que colaboradores externos aprovem solicitações pull neste repositório.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Em quanto tempo receberei uma resposta sobre minha solicitação de alteração?

Normalmente, as solicitações de pull são revisadas dentro de 10 dias úteis.

## <a name="more-resources"></a>Mais recursos

* Para saber mais sobre como usar o Git e GitHub, confira primeiro o [GitHub Ajuda].

[Ajuda do GitHub]: https://docs.github.com
[Configurar o Git]: https://docs.github.com/get-started/quickstart/set-up-git
