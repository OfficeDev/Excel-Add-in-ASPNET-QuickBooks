<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Close.aspx.cs" Inherits="QbAdd_inDotNetWeb.Close" %>

<!DOCTYPE html>
<!--
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Close</title>
    <script src="//appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js" type="text/javascript"></script>
    <script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>

    <script>
        Office.initialize = function (reason) {
            var message = {};
            message.token = "<%=HttpContext.Current.Session["accessToken"].ToString()%>";
            message.secret = "<%=HttpContext.Current.Session["accessTokenSecret"].ToString()%>";
            message.realm = "<%=HttpContext.Current.Session["realm"].ToString()%>";
            Office.context.ui.messageParent(JSON.stringify(message));
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
    </form>
</body>
</html>
