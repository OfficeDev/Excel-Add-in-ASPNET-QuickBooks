# <a name="excel-add-in-with-aspnet-and-quickbooks"></a>Excel 增益集與 ASP.NET 和 QuickBooks

您的 Excel 增益集可以連接到像是 QuickBooks 的服務，並且將資料匯入至 Excel 試算表。這個 Excel 增益集將示範如何連接至 QuickBooks、從 QuickBooks 提供的沙箱帳戶取得範例費用資料，**Sandbox Company_US_1**，並且將範例資料匯入至試算表。增益集也提供按鈕以根據範例資料建立圖表。

## <a name="table-of-contents"></a>目錄

* [必要條件](#prerequisites)
* [設定專案](#configure-the-project)
* [執行專案](#run-the-project)
* [瞭解程式碼](#understand-the-code)
* [連線至 Office 365](#connect-to-office-365)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## <a name="prerequisites"></a>必要條件

* [QuickBooks 開發人員](https://developer.intuit.com/)帳戶
* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Visual Studio 的 Office 開發人員工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

## <a name="configure-the-project"></a>設定專案

在 developer.intuit.com 設定您的應用程式以開始使用。

1. 移至 https://developer.intuit.com/ 並且註冊開發人員帳戶，然後登入。
2. 在右上角，選擇 **[我的應用程式]**，並選取應用程式，或者按一下 **[建立新應用程式]**。 
3. 一旦選取應用程式，選擇 **[開發** | **金鑰]**，並複製 **[OAuth 家庭用戶金鑰]** 和 **[OAuth 家庭用戶密碼]**，將其放置在您稍後可以存取的位置。
4. 下載或複製範例到本機電腦。
5. 在 Visual Studio 中開啟解決方案檔案 **QbAdd-inDotNet.sln**。
6. 在 Visual Studio 中，開啟 **Web.config** 並插入 `ConsumerKey` 和 `ConsumerSecret` 的值，如下所示。

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

## <a name="run-the-project"></a>執行專案

1. 按下 F5 以執行專案。

2. 從 Excel 中的功能區選取指令按鈕，來啟動增益集。<br>![QuickBooks Excel 增益集命令按鈕](../readme-images/readme_command_image.PNG)  

3. 按一下 **[連線到 QuickBooks]**，來啟動 QuickBooks 的登入視窗。<br>![工作窗格登入](../readme-images/readme_image_taskpane.PNG)

4. 如果 Visual Studio 中有錯誤視窗開啟，請按一下 **[繼續]**，並且瀏覽回 Excel。這個錯誤與範例不相關。<br>![Visual Studio 錯誤視窗](../readme-images/readme_image_error.PNG)

5. 使用您的 QuickBooks 開發人員帳戶來登入 QuickBooks。<br>![QuickBooks 登入對話視窗](../readme-images/readme_image_signin.PNG)

6. 按一下 **[授權]**，以便讓 QuickBooks 將資料傳送至增益集。<br>![QuickBooks 授權對話視窗](../readme-images/readme_image_authorize.PNG) <br> 工作窗格將會顯示兩個動作以供選擇。 <br>選取 [動作工作窗格]![](../readme-images/readme_image_action.PNG)

8. 選擇 **[取得費用]**，即可將費用從 QuickBooks 匯入到試算表。 <br>![費用試算表](../readme-images/readme_image_expenses.PNG)

9. 選擇 **[建立圖表]**，即可插入圖表。 <br>![插入圖表](../readme-images/readme_image_chart.PNG)

## <a name="understand-the-code"></a>瞭解程式碼

* [Home.html](QbAdd-inDotNetWeb/Home.html) - 定義啟動時，以及使用者登入之後的工作窗格頁面。
* [Home.js](QbAdd-inDotNetWeb/Home.js) - 處理登入、登出、取得費用以及插入圖表的使用者互動。在這裡，會呼叫 `dialogDisplayAsync` API 以開啟對話方塊視窗，讓使用者登入 QuickBooks。
* [QbAdd-inDotNet.xml](QbAdd-inDotNet/QbAdd-inDotNetManifest/QbAdd-inDotNet.xml) - 增益集的資訊清單檔。 
* [QuickBooksController.cs](QbAdd-inDotNetWeb/Controllers/QuickBooksController.cs) - 從 QuickBooks 取得費用資料。
* [FunctionFile.js](QbAdd-inDotNetWeb/Functions/FunctionFile.js) - 將圖表新增至 Excel。
* [OAuthManager.aspx.cs](QbAdd-inDotNetWeb/OAuthManager.aspx.cs) - 從對話方塊 API 處理登入至 QuickBooks。

## <a name="questions-and-comments"></a>問題和建議

我們樂於獲得您關於 *Excel 增益集與 ASPNET 和 QuickBooks* 範例的意見反應。您可以在此儲存機制的 [問題]** 區段中，將您的意見反應傳送給我們。請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關 Office 365 開發的一般問題。務必以 [Office365] 和 [API] 標記您的問題。

## <a name="additional-resources"></a>其他資源

* [Office 365 API 文件](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API 工具](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office 開發人員中心](http://dev.office.com/)
* [Office 365 API 入門專案和程式碼範例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft.著作權所有，並保留一切權利。


此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
