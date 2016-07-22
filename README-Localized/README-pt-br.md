# Suplemento do Excel com o ASP.NET e QuickBooks

O Suplemento do Excel pode se conectar a um serviço, como o QuickBooks, e importar dados para sua planilha do Excel. Esse Suplemento do Excel demonstra como se conectar ao QuickBooks, obtém dados de exemplos de despesas de uma conta de área restrita fornecidos pelo QuickBooks, **Área Restrita Company_US_1**, e importa os dados de exemplo para uma planilha. O suplemento também fornece um botão para criar um gráfico usando os dados de exemplo.

## Sumário

* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Executar o projeto](#run-the-project)
* [Compreender o código](#understand-the-code)
* [Conectar-se ao Office 365](#connect-to-office-365)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## Pré-requisitos

* Uma conta de [desenvolvedor do QuickBooks](https://developer.intuit.com/)
* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools for Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

## Configurar o projeto

Configure seu aplicativo na página developer.intuit.com para começar.

1. Acesse https://developer.intuit.com/, inscreva-se para uma conta de desenvolvedor e entre.
2. No canto superior direito, escolha **Meus Aplicativos** e escolha um aplicativo ou clique em **Criar novo aplicativo**. 
3. Assim que o aplicativo for selecionado, escolha **Chaves de**|**Desenvolvimento** e copie a **Chave do Consumidor OAuth** e o **Segredo do Consumidor OAuth** para um local que possa ser acessado posteriormente.
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

## Executar o projeto

1. Pressione F5 para executar o projeto.

2. Inicie o suplemento selecionando o botão de comando da faixa de opções do Excel.<br><img src="readme-images/readme_command_image.PNG" alt="Botão de comando do Suplemento do Excel do QuickBooks"></img>  

3. Clique em **Conectar-se ao QuickBooks** para iniciar a janela de início de sessão do QuickBooks.<br><img src="readme-images/readme_image_taskpane.PNG" alt="Entrada no painel de tarefas"></img>

4. Se uma janela de erro for exibida no Visual Studio, clique em **Continuar** e navegue para o Excel. Esse erro está relacionado aos exemplos.<br><img src="readme-images/readme_image_error.PNG" alt="Janela de erro do Visual Studio"></img>

5. Entrar no QuickBooks com sua conta de desenvolvedor do QuickBooks.<br><img src="readme-images/readme_image_signin.PNG" alt="Janela de diálogo de entrada do QuickBooks"></img>

6. Clique em **Autorizar** para permitir que o QuickBooks envie dados para o suplemento<br><img src="readme-images/readme_image_authorize.PNG" alt="Janela de diálogo de autorização do QuickBooks"></img> <br> O painel de tarefas exibirá duas ações à sua escolha. <br><img src="readme-images/readme_image_action.PNG" alt="Painel de tarefas para escolher ações"></img>

8. Escolha **Obter despesas** para importar as despesas do QuickBooks para uma planilha. <br><img src="readme-images/readme_image_expenses.PNG" alt="Planilha de despesas"></img>

9. Escolha **Criar Gráfico** para inserir um gráfico. <br><img src="readme-images/readme_image_chart.PNG" alt="Inserir gráfico"></img>

## Compreender o código

* [Home.HTML](QbAdd-inDotNetWeb/home.html) - define a página do painel de tarefas na inicialização e depois que o usuário tiver efetuado logon.
* [Home.js](QbAdd-inDotNetWeb/home.js) - trata da interação do usuário para entrar, sair, obter despesas e inserir gráfico. Aqui, a API `dialogDisplayAsync` é chamada para abrir uma janela de diálogo para o usuário entrar no QuickBooks.
* [QbAdd inDotNet.xml](QbAdd-inDotNet/QbAdd-inDotNetManifest/QbAdd-inDotNet.xml) - o arquivo de manifesto do suplemento. 
* [QuickBooksController.cs](QbAdd-inDotNetWeb/Controllers/QuickBooksController.cs) - obtém dados de despesas a partir do QuickBooks.
* [FunctionFile.js](QbAdd-inDotNetWeb/Functions/FunctionFile.js) - adiciona um gráfico ao Excel.
* [OAuthManager.aspx.cs](QbAdd-inDotNetWeb/OAuthManager.aspx.cs) - trata das entradas no QuickBooks realizadas a partir da API de diálogo.

## Perguntas e comentários

Adoraríamos receber seus comentários sobre o exemplo do *Suplemento do Excel com o ASPNET e QuickBooks*. Você pode enviar comentários na seção *Problemas* deste repositório. As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Não deixe de marcar as perguntas com [Office365] e [API].

## Recursos adicionais

* [Documentação de APIs do Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Ferramentas de API do Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* [Exemplos de código e projetos iniciais de APIs do Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Direitos autorais
Copyright © 2016 Microsoft. Todos os direitos reservados.

