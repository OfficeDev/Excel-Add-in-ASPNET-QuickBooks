# <a name="excel-add-in-with-aspnet-and-quickbooks"></a>Suplemento do Excel com o ASP.NET e QuickBooks

O Suplemento do Excel pode se conectar a um serviço, como o QuickBooks, e importar dados para sua planilha do Excel. Esse Suplemento do Excel demonstra como se conectar ao QuickBooks, obtém dados de exemplos de despesas de uma conta de área restrita fornecidos pelo QuickBooks, **Área Restrita Company_US_1**, e importa os dados de exemplo para uma planilha. O suplemento também fornece um botão para criar um gráfico usando os dados de exemplo.

## <a name="table-of-contents"></a>Sumário

* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Executar o projeto](#run-the-project)
* [Compreender o código](#understand-the-code)
* [Conectar-se ao Office 365](#connect-to-office-365)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="prerequisites"></a>Pré-requisitos

* Uma conta de [desenvolvedor do QuickBooks](https://developer.intuit.com/)
* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools para Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

## <a name="configure-the-project"></a>Configurar o projeto

Configure seu aplicativo na página developer.intuit.com para começar.

1. Acesse https://developer.intuit.com/, inscreva-se para uma conta de desenvolvedor e entre.
2. No canto superior direito, escolha **Meus Aplicativos** e escolha um aplicativo ou clique em **Criar novo aplicativo**. 
3. Assim que o aplicativo for selecionado, escolha **Chaves de** | **Desenvolvimento** e copie a **Chave do Consumidor OAuth** e o **Segredo do Consumidor OAuth** para um local onde elas possam ser acessadas posteriormente.
4. Baixe ou clone o exemplo para o computador local.
5. Abra o arquivo de solução **QbAdd inDotNet.sln** no Visual Studio.
6. No Visual Studio, abra **Web.config** e insira os valores para `ConsumerKey` e `ConsumerSecret`, dessa forma.

```
<appSettings>
    <!-- QuickBooks Settings -->
    <add key="ConsumerKey" value="insert your OAuth Consumer Key here" />
    <add key="ConsumerSecret" value="insert your OAuth Consumer Secret here" />
    <add key="OauthLink" value="https://oauth.intuit.com/oauth/v1" />
    <add key="AuthorizeUrl" value="https://workplace.intuit.com/Connect/Begin" />
    <add key="RequestTokenUrl" value="https://oauth.intuit.com/oauth/v1/get_request_token" />
    <add key="AccessTokenUrl" value="https://oauth.intuit.com/oauth/v1/get_access_token" />
    <add key="ServiceContext.BaseUrl.Qbo" value="https://sandbox-quickbooks.api.intuit.com/" />
    <add key="DeepLink" value="sandbox.qbo.intuit.com" />
  </appSettings>
```

## <a name="run-the-project"></a>Executar o projeto

1. Pressione F5 para executar o projeto.

2. Inicie o suplemento selecionando o botão de comando da faixa de opções no Excel<br>![Botão de comando do Suplemento QuickBooks do Excel](../readme-images/readme_command_image.PNG)  

3. Clique em **Conectar-se ao QuickBooks** para iniciar a janela de entrada do QuickBooks.<br>![Entrada do painel de tarefas](../readme-images/readme_image_taskpane.PNG)

4. Se uma janela de erro for exibida no Visual Studio, clique em **Continuar** e navegue para o Excel. Este erro não está relacionado ao exemplo.<br>![Janela de erro do Visual Studio](../readme-images/readme_image_error.PNG)

5. Entre no QuickBooks com sua senha de desenvolvedor do QuickBooks.<br>![Janela da caixa de diálogo de entrada do QuickBooks](../readme-images/readme_image_signin.PNG)

6. Clique em **Autorizar** para permitir que o QuickBooks envie dados ao suplemento.<br>![Janela da caixa de diálogo de autorização do QuickBooks](../readme-images/readme_image_authorize.PNG) <br> O painel de tarefas exibirá duas ações à sua escolha. <br>![Selecione o painel de tarefas de ação](../readme-images/readme_image_action.PNG)

8. Escolha **Obter despesas** para importar as despesas do QuickBooks para uma planilha. <br>![Planilhas de despesas](../readme-images/readme_image_expenses.PNG)

9. Escolha **Criar Gráfico** para inserir um gráfico. <br>![Inserir gráfico](../readme-images/readme_image_chart.PNG)

## <a name="understand-the-code"></a>Compreender o código

* [Home.HTML](QbAdd-inDotNetWeb/Home.html) – define a página do painel de tarefas na inicialização e depois que o usuário tiver feito logon.
* [Home.js](QbAdd-inDotNetWeb/Home.js) – trata da interação do usuário para entrar, sair, obter despesas e inserir gráfico. Aqui, a API `dialogDisplayAsync` é chamada para abrir uma janela de diálogo para o usuário entrar no QuickBooks.
* [QbAdd-inDotNet.xml](QbAdd-inDotNet/QbAdd-inDotNetManifest/QbAdd-inDotNet.xml) – o arquivo de manifesto do suplemento. 
* [QuickBooksController.cs](QbAdd-inDotNetWeb/Controllers/QuickBooksController.cs) – obtém dados de despesas a partir do QuickBooks.
* [FunctionFile.js](QbAdd-inDotNetWeb/Functions/FunctionFile.js) – adiciona um gráfico ao Excel.
* [OAuthManager.aspx.cs](QbAdd-inDotNetWeb/OAuthManager.aspx.cs) – trata das entradas no QuickBooks realizadas a partir da API de diálogo.

## <a name="questions-and-comments"></a>Perguntas e comentários

Adoraríamos receber seus comentários sobre o exemplo do *Suplemento do Excel com o ASPNET e QuickBooks*. Você pode enviar comentários na seção *Problemas* deste repositório. As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Não deixe de marcar as perguntas com [Office365] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação de APIs do Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Ferramentas de API do Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* [Exemplos de código e projetos iniciais de APIs do Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright © 2016 Microsoft. Todos os direitos reservados.


Este projeto adotou o [Código de Conduta de Software Livre da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
