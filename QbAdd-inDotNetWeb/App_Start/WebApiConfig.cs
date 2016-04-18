/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.SessionState;
using System.Web.Routing;
using System.Web;
using System.Web.Http.WebHost;

namespace QbAdd_inDotNetWeb
{
    public static class WebApiConfig
    {
        /// <summary>
        /// Web API configuration and service
        /// </summary>
        /// <param name="config"></param>
        public static void Register(HttpConfiguration config)
        {
            // Web API routes
            config.MapHttpAttributeRoutes();

            // Routes the request to the action 
            config.Routes.MapHttpRoute(
                name: "ActionApi",
                routeTemplate: "api/{action}/{id}",
                defaults: new { id = RouteParameter.Optional, controller = "QuickBooks"}
            );

        }
    }
}
