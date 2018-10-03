# <a name="identity-api-requirement-sets"></a>Identificar conjuntos de requisitos da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Suplementos do Office usam conjuntos de requisito especificados no manifesto ou uma seleção de tempo de execução para determinar se um host do Office oferece suporte a APIs que precisa de um suplemento. Para obter mais informações, consulte [define versões do Office e o requisito](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Suplementos do Office enfrentar diversas versões do Office. A tabela a seguir lista os conjuntos de requisito de API de identidade, os aplicativos de host do Office que oferecem suporte a números de conjunto e a compilação ou a versão esse requisito para o aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 365 para Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com e Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/D | Visualizar **& #42;** | Em breve | Visualizar **& #42;**| Disponível | Disponível| Em breve | Em breve |

> **& #42;** Durante a fase de visualização, a API de identidade é suportada no Windows 2016 e Mac somente para usuários no programa internas usando a opção Fast. Para ingressar no programa internas, consulte a [ser um Insider do Office](https://products.office.com/office-insider?tab=tab-1). Para alternar para o Fast track, consulte [Insider Fast](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos de API comum

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

O IdentityAPI 1.1 de Logon Único é a primeira versão da API. Para obter detalhes sobre essa API, consulte a seção de [referência da API de SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) de [Habilitar o SSO em um suplemento](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar hosts do Office e requisitos da API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
