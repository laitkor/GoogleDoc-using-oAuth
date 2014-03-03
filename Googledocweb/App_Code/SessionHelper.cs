using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Googledocweb
{
    class SessionHelper
    {
        private static string _Token = "Token";
        private static string _accessToken = "AccessToken";
        private static string _AuthenticationError = "AuthenticationError";

        public static string Token
        {
            get
            {
                Object token = getSession(_Token);
                if (token != null)
                {
                    return token.ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                setSession(_Token, value);
            }
        }

        public static string AccessToken
        {
            get
            {
                Object accesstoken = getSession(_accessToken);
                if (accesstoken != null)
                {
                    return accesstoken.ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                setSession(_accessToken, value);
            }
        }

        public static string AuthenticationError
        {
            get
            {
                Object authenticationError = getSession(_AuthenticationError);
                if (authenticationError != null)
                {
                    return authenticationError.ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                setSession(_AuthenticationError, value);
            }
        }

        private static void setSession(string name, Object value)
        {
            HttpContext.Current.Session[name] = value;
        }

        private static Object getSession(string name)
        {
            Object toret = null;
            if (HttpContext.Current.Session[name] != null)
            {
                try
                {
                    toret = HttpContext.Current.Session[name];
                }
                catch (Exception ex)
                {
                }
            }
            return toret;
        } 
    }
}
