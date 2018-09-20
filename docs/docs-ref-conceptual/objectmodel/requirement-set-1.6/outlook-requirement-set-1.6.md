# <a name="outlook-add-in-api-requirement-set-16"></a>Requisito de Outlook suplemento API definir 1.6

O subconjunto de APIs de suplemento do Outlook para as APIs de JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

## <a name="whats-new-in-16"></a>What's new in 1.6?

Set-requisito 1.6 inclui todos os recursos do [conjunto de requisito de 1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Ele adicionado os recursos a seguir.

- Adicionadas novas APIs para suplementos contextuais para obter a entidade ou RegEx corresponde o usuário selecionado para ativar o suplemento.
- Adicionada uma nova API para abrir um novo formulário de mensagem.
- Adicionada a capacidade para o suplemento determinar o tipo de conta da caixa de correio do usuário.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada um usuário tiver selecionado. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): adiciona uma nova função que retorna os valores de cadeia de caracteres em uma correspondência realçada que coincidem com as expressões regulares definidas no arquivo XML de manifesto. As correspondências realçadas aplicam-se aos suplementos contextuais.
- Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): adiciona uma nova função que abre um novo formulário de mensagem.
- Adicionado [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): adiciona um novo membro para o perfil de usuário que indica o tipo de conta do usuário.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)