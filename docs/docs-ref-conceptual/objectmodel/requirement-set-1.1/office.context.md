
# <a name="context"></a>context

### [Office](Office.md). context

O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).


##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

### <a name="namespaces"></a>Namespaces

[caixa de correio](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e Microsoft Outlook na web.

### <a name="members"></a>Membros

####  <a name="displaylanguage-string"></a>displayLanguage :String

Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.

O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

##### <a name="example"></a>Exemplo

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.

O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.

##### <a name="type"></a>Tipo:

*   [RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|