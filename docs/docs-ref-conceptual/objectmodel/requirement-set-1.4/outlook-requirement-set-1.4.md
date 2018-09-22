# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.4

O Outlook suplemento subconjunto de API da API do JavaScript para Office inclui objetos, métodos, propriedades e suplemento de eventos que você pode usar no Outlook.

> [!NOTE]
> Esta documentação é para um [requisito definido](/javascript/office/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisito mais recente.

## <a name="whats-new-in-14"></a>Novidades na versão 1.4

O conjunto de requisitos 1.4 inclui todos os recursos do [Conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Exibe uma caixa de diálogo em um host do Office.
- Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.
- Objeto de [diálogo](/javascript/api/office/office.dialog) adicionado: O objeto retornado quando o [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) método for chamado.

## <a name="see-also"></a>Confira também

- 
  [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)