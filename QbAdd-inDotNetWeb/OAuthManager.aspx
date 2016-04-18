<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OAuthManager.aspx.cs" Inherits="QbAdd_inDotNetWeb.OAuthManager" %>

<!DOCTYPE html>
<!--
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <%if (null != HttpContext.Current.Session["accessToken"]) { %>
                <script>
                    window.location.href = 'https://localhost:44300/close.aspx';
                </script>
            <% } %>
        </div>
    </form>
</body>
</html>
