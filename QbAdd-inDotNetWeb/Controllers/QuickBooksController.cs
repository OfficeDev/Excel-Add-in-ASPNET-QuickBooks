/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using System.Web;
using System.Configuration;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.Security;


namespace QbAdd_inDotNetWeb
{
    /// <summary>
    /// Controller which connects to QuickBooks and gets expenses.
    /// This flow will make use of Data Service SDK V2 to create an OAuthRequest and connect
    /// to Customer Data under the service context and store data in a list.
    /// </summary>
    public class QuickBooksController : ApiController
    {

        private String realmId, accessToken, accessTokenSecret, consumerKey, consumerSecret;

        /// <summary>
        /// Within an OAuth session, pass a token to QuickBooks, and issue a query to make a GET call
        /// for Customer expenses.
        /// Puts the data returned from the service into an array.
        /// </summary>
        /// <param name="n"></param>
        /// <returns>A collection of Purchases</returns>
        [HttpGet]
        public IEnumerable<Purchase> GetExpenses(int n)
        {
            consumerKey = ConfigurationManager.AppSettings["consumerKey"].ToString();
            consumerSecret = ConfigurationManager.AppSettings["consumerSecret"].ToString();

            realmId = HttpContext.Current.Session["realm"].ToString();
            accessToken = HttpContext.Current.Session["accessToken"].ToString();
            accessTokenSecret = HttpContext.Current.Session["accessTokenSecret"].ToString();

            IntuitServicesType intuitServicesType = IntuitServicesType.QBO;
            OAuthRequestValidator oauthValidator = new OAuthRequestValidator(accessToken, accessTokenSecret, consumerKey, consumerSecret);
            ServiceContext context = new ServiceContext(realmId, intuitServicesType, oauthValidator);
            context.IppConfiguration.BaseUrl.Qbo = ConfigurationManager.AppSettings["ServiceContext.BaseUrl.Qbo"].ToString();

            DataService dataService = new DataService(context);
            List<Purchase> expenses = dataService.FindAll(new Purchase(), 1, n).ToList();
            return expenses;
        }

        /// <summary>
        /// Helper method to set token values for the current HTTP session.
        /// </summary>
        /// <param name="token"></param>
        /// <param name="secret"></param>
        /// <param name="realm"></param>
        /// <returns>An HTTP response message</returns>
        [HttpGet]
        public HttpResponseMessage SetToken(string token, string secret, string realm)
        {
            HttpContext.Current.Session["accessToken"] = token;
            HttpContext.Current.Session["accessTokenSecret"] = secret;
            HttpContext.Current.Session["realm"] = realm;

            return Request.CreateResponse(HttpStatusCode.OK, "Success");     
        }

        /// <summary>
        /// Helper method to get the token values for the current HTTP session.
        /// </summary>
        /// <returns>An HTTP response message</returns>
        public HttpResponseMessage GetToken()
        {
            HttpStatusCode code = HttpStatusCode.NotFound;
            string message = "NotFound";
            if (null != HttpContext.Current.Session["accessToken"] && "" != HttpContext.Current.Session["accessToken"].ToString())
            {
                code = HttpStatusCode.OK;
                message = "Success";
            }

            return Request.CreateResponse(code, message);
        }

        /// <summary>
        /// Helper method to clear cached token for the current session.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public HttpResponseMessage ClearToken()
        {
            HttpContext.Current.Session["accessToken"] = "";
            HttpContext.Current.Session["accessTokenSecret"] = "";

            return Request.CreateResponse(HttpStatusCode.OK, "Success");
        }
    }
}