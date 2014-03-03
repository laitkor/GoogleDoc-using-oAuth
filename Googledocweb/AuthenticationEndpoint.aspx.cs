using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Security;
using System.Web.Security;
using DotNetOpenAuth.OpenId.RelyingParty;
using DotNetOpenAuth.OpenId.Extensions.AttributeExchange;

namespace Googledocweb
{
    public partial class AuthenticationEndpoint : System.Web.UI.Page
    {

        
        private const string CALLBACK_PARAMETER = "callback";
        private const string RETURNURL_PARAMETER = "ReturnUrl";
        private const string AUTHENTICATION_ENDPOINT =
                                "~/AuthenticationEndpoint.aspx";
        private const string GOOGLE_OAUTH_ENDPOINT =
            "https://www.google.com/accounts/o8/id";
        /// <summary>
        /// Call Google Authentication
        /// </summary>
        protected void Page_Load(object sender, EventArgs e)
        {
            if (User != null && User.Identity != null && User.Identity.IsAuthenticated
                && !String.IsNullOrWhiteSpace(Request.Params[RETURNURL_PARAMETER]))
            {
                Response.Redirect(Request.Params[RETURNURL_PARAMETER]);
            }
            else
            {
                //Check if either to handle a call back or start an authentication
                if (Request.Params[CALLBACK_PARAMETER] == "true")
                {
                    HandleAuthenticationCallback(); //Google has performed a callback, let's analyze it
                   
                }
                else
                {
                   
                        PerformGoogleAuthentication();  //There is no callback parameter, 
                        //so it looks like we want to sign in our fellow
                   
                }
            }
 
        }



       
        protected void PerformGoogleAuthentication()
        {
            using (OpenIdRelyingParty openid = new OpenIdRelyingParty())
            {
                //Set up the callback URL
                Uri callbackUrl = new Uri(
                    String.Format("{0}{1}{2}{3}?{4}=true",
                    (Request.IsSecureConnection) ? "https://" : "http://",
                    Request.Url.Host,
                    (Request.Url.IsDefaultPort) ?
                        String.Empty : String.Concat(":", Request.Url.Port),
                    Page.ResolveUrl(AUTHENTICATION_ENDPOINT),
                    CALLBACK_PARAMETER
                    ));

                //Set up request object for Google Authentication
                IAuthenticationRequest request =
                    openid.CreateRequest(GOOGLE_OAUTH_ENDPOINT,
                    DotNetOpenAuth.OpenId.Realm.AutoDetect, callbackUrl);


                //Let's tell Google, what we want to have from the user:
                var fetch = new FetchRequest();
                fetch.Attributes.AddRequired(WellKnownAttributes.Contact.Email);
                fetch.Attributes.AddRequired(WellKnownAttributes.Name.First);
                fetch.Attributes.AddRequired(WellKnownAttributes.Name.Last);
                request.AddExtension(fetch);

                //Redirect to Google Authentication
                request.RedirectToProvider();
            }
        }

        /// <summary>
        /// Handle the response that Google posted back
        /// </summary>
        public void HandleAuthenticationCallback()
        {
            
            OpenIdRelyingParty openid = new OpenIdRelyingParty();
            var response = openid.GetResponse();
            if (response == null) { ThrowSecurityException(); return; }

            switch (response.Status)
            {
                case AuthenticationStatus.Authenticated:
                    var fetch = response.GetExtension<FetchResponse>();
                    string email = string.Empty;
                    string firstname = string.Empty;
                    string lastname = string.Empty;
                    string strToken = "";
                    ;
                    if (fetch != null)
                    {
                        SessionHelper.AuthenticationError = null;
                        email = fetch.GetAttributeValue(WellKnownAttributes.Contact.Email);
                        firstname = fetch.GetAttributeValue(WellKnownAttributes.Name.First);
                        lastname = fetch.GetAttributeValue(WellKnownAttributes.Name.Last);
                       // strToken = Request.QueryString["oauth_token"].ToString();
                        //Response.Redirect("~/AuthenticationEndpoint.aspx");
                        FormsAuthentication.SetAuthCookie(response.ClaimedIdentifier, false);
                        FormsAuthentication.RedirectFromLoginPage(response.ClaimedIdentifier, false);
                    }
                    else //we didn't fetch any info. Too bad.
                    {
                        ThrowSecurityException();
                    }
                    break;
                case AuthenticationStatus.Canceled:
                    SessionHelper.AuthenticationError = "access_denied";
                    UtilityCode.Setmessage("Authentication failed for accessing Google Doc's. Click <a href='AuthenticationEndPoint.aspx'>here</a> to authenticate and authorize access to Google Doc's.", Lbl_Msg, MessageType.Warning);
                    break;
                //You might want to differ the states a bit more
                default:
                    ThrowSecurityException();
                    break;
            }
        }

        /// <summary>
        /// This exception throws a simple and ugly error message. 
        /// You may improve this message ;-)
        /// </summary>
        public void ThrowSecurityException()
        {
            throw new SecurityException("Authentication failed");
        }
 
    }
}