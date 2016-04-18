/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Configuration;
using DevDefined.OAuth.Consumer;
using DevDefined.OAuth.Framework;


namespace QbAdd_inDotNetWeb
{
    public partial class OAuthManager : System.Web.UI.Page
    {
        /// <summary>
        /// Stores configuration settings for OAuth.
        /// </summary>
        private string requestTokenUrl = ConfigurationManager.AppSettings["RequestTokenUrl"];
        private string accessTokenUrl = ConfigurationManager.AppSettings["AccessTokenUrl"];
        private string authorizeUrl = ConfigurationManager.AppSettings["AuthorizeUrl"];
        private string oauthUrl = ConfigurationManager.AppSettings["OauthLink"];
        private string consumerKey = ConfigurationManager.AppSettings["ConsumerKey"];
        private string consumerSecret = ConfigurationManager.AppSettings["ConsumerSecret"];
        private string oauthCallbackUrl = "https://localhost:44300/OauthManager.aspx";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString.Count > 0)
            {
                List<string> queryKeys = new List<string>(Request.QueryString.AllKeys);
                if (queryKeys.Contains("connect"))
                {
                    FireAuth();
                }
                if (queryKeys.Contains("oauth_token"))
                {
                    ReadToken();
                }
            }

        }

        /// <summary>
        /// Starts the authentication process.
        /// </summary>
        private void FireAuth()
        {

            IOAuthSession session = CreateSession();
            IToken requestToken = session.GetRequestToken();
            HttpContext.Current.Session["requestToken"] = requestToken;
            var authUrl = string.Format("{0}?oauth_token={1}&oauth_callback={2}", authorizeUrl, requestToken.Token, UriUtility.UrlEncode(oauthCallbackUrl));
            HttpContext.Current.Session["oauthLink"] = authUrl;

            HttpContext.Current.Response.Redirect(authUrl);
        }

        /// <summary>
        /// Gets the token for the current session.
        /// </summary>
        private void ReadToken()
        {
            HttpContext.Current.Session["oauthToken"] = Request.QueryString["oauth_token"].ToString(); ;
            HttpContext.Current.Session["oauthVerifyer"] = Request.QueryString["oauth_verifier"].ToString();
            HttpContext.Current.Session["realm"] = Request.QueryString["realmId"].ToString();
            HttpContext.Current.Session["dataSource"] = Request.QueryString["dataSource"].ToString();
            //Stored in a session for demo purposes.
            //Production applications should securely store the access token.
            IOAuthSession clientSession = CreateSession();
            IToken accessToken = clientSession.ExchangeRequestTokenForAccessToken((IToken)HttpContext.Current.Session["requestToken"], HttpContext.Current.Session["oauthVerifyer"].ToString());
            HttpContext.Current.Session["accessToken"] = accessToken.Token;
            HttpContext.Current.Session["accessTokenSecret"] = accessToken.TokenSecret;

        }

        /// <summary>
        /// Creates a new OAuth session.
        /// </summary>
        /// <returns></returns>
        protected IOAuthSession CreateSession()
        {
            var consumerContext = new OAuthConsumerContext
            {
                ConsumerKey = consumerKey,
                ConsumerSecret = consumerSecret,
                SignatureMethod = SignatureMethod.HmacSha1
            };
            return new OAuthSession(consumerContext,
                                    requestTokenUrl,
                                    oauthUrl,
                                    accessTokenUrl);
        }



    }
}