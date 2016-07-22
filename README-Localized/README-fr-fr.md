# Complément Excel avec ASP.NET et QuickBooks

Votre complément Excel peut se connecter à un service tel que QuickBooks et importer des données dans votre feuille de calcul Excel. Ce complément Excel indique comment se connecter à QuickBooks, obtient des exemples de données de dépenses auprès d’un compte sandbox fourni par QuickBooks, **Sandbox Company_US_1**, et importe les exemples de données dans une feuille de calcul. Le complément fournit également un bouton permettant de créer un graphique à partir des exemples de données.

## Table des matières

* [Conditions requises](#prerequisites)
* [Configurer le projet](#configure-the-project)
* [Exécuter le projet](#run-the-project)
* [Comprendre le code](#understand-the-code)
* [Se connecter à Office 365](#connect-to-office-365)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## Conditions requises

* Un compte de développeur [QuickBooks](https://developer.intuit.com/)
* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Outils de développement Office pour Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

## Configurer le projet

Pour commencer, configurez votre application à l’adresse developer.intuit.com.

1. Accédez à https://developer.intuit.com/, inscrivez-vous pour obtenir un compte de développeur, puis connectez-vous.
2. Dans le coin supérieur droit, sélectionnez **Mes applications**, puis choisissez une application ou cliquez sur **Créer une application**. 
3. Après avoir sélectionné l’application, choisissez **Développement**|**Clés**, puis copiez la **clé de consommateur OAuth** et le **secret de consommateur OAuth** à un emplacement auquel vous pourrez accéder ultérieurement.
4. Téléchargez ou clonez l’exemple sur votre ordinateur local.
5. Ouvrez le fichier de solution **QbAdd-inDotNet.sln** dans Visual Studio.
6. Dans Visual Studio, ouvrez **Web.config** et insérez les valeurs de `ConsumerKey` et `ConsumerSecret`, comme suit.

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

## Exécuter le projet

1. Appuyez sur F5 pour exécuter le projet.

2. Lancez le complément en cliquant sur le bouton de commande du ruban dans Excel.<br><img src="readme-images/readme_command_image.PNG" alt="Bouton de commande de complément Excel QuickBooks"></img>  

3. Cliquez sur **Se connecter à QuickBooks** pour lancer la fenêtre de connexion QuickBooks.<br><img src="readme-images/readme_image_taskpane.PNG" alt="Volet Office - Connexion"></img>

4. Si une fenêtre d’erreur s’ouvre dans Visual Studio, cliquez sur **Continuer** et retournez dans Excel. Cette erreur n’est pas liée à l’exemple.<br><img src="readme-images/readme_image_error.PNG" alt="Fenêtre d’erreur Visual Studio"></img>

5. Connectez-vous à QuickBooks avec votre compte de développeur QuickBooks.<br><img src="readme-images/readme_image_signin.PNG" alt="Boîte de dialogue de connexion à QuickBooks"></img>

6. Cliquez sur **Autoriser** pour autoriser QuickBooks à envoyer des données au complément.<br><img src="readme-images/readme_image_authorize.PNG" alt="Boîte de dialogue d’autorisation QuickBooks"></img> <br> Le volet Office affiche deux actions entre lesquelles choisir.<br><img src="readme-images/readme_image_action.PNG" alt="Volet Office de sélection d’action"></img>

8. Choisissez **Obtenir des dépenses** pour importer des dépenses à partir de QuickBooks vers une feuille de calcul. <br><img src="readme-images/readme_image_expenses.PNG" alt="Feuille de calcul de dépenses"></img>

9. Choisissez **Créer un graphique** pour insérer une graphique. <br><img src="readme-images/readme_image_chart.PNG" alt="Insérer un graphique"></img>

## Comprendre le code

* [Home.HTML](QbAdd-inDotNetWeb/home.html) : définit la page de volet Office lors du démarrage et après la connexion de l’utilisateur.
* [Home.js](QbAdd-inDotNetWeb/home.js) : gère les interactions de l’utilisateur pour la connexion, la déconnexion, l’obtention des dépenses et l’insertion de graphiques. Dans ce cas, l’API `dialogDisplayAsync` est appelée pour ouvrir une boîte de dialogue afin que l’utilisateur se connecte à QuickBooks.
* [QbAdd-inDotNet.xml](QbAdd-inDotNet/QbAdd-inDotNetManifest/QbAdd-inDotNet.xml) : fichier manifeste pour le complément. 
* [QuickBooksController.cs](QbAdd-inDotNetWeb/Controllers/QuickBooksController.cs) : récupère les données de dépenses auprès de QuickBooks.
* [FunctionFile.js](QbAdd-inDotNetWeb/Functions/FunctionFile.js) : ajoute un graphique dans Excel.
* [OAuthManager.aspx.cs](QbAdd-inDotNetWeb/OAuthManager.aspx.cs) : gère la connexion à QuickBooks à partir de l’API de boîte de dialogue.

## Questions et commentaires

Nous aimerions recevoir vos commentaires sur l’exemple *Complément Excel avec ASP.NET et QuickBooks*. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel. Si vous avez des questions sur le développement d’Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Veillez à poser vos questions avec les balises [API] et [Office365].

## Ressources supplémentaires

* [Documentation sur les API Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Outils d’API Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centre de développement Office](http://dev.office.com/)
* [Projets de démarrage et exemples de code des API Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft. Tous droits réservés.

