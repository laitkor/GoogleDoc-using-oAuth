using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Googledocweb
{
    public partial class oauth2callback : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string Token = Request.QueryString["code"];
            Session.Clear();
            if (!string.IsNullOrWhiteSpace(Token))
            {
                SessionHelper.Token = Token;
                SessionHelper.AuthenticationError = null;
            }

            string authenticationError = Request.QueryString["error"];
            if (!string.IsNullOrWhiteSpace(authenticationError))
            {
                SessionHelper.AuthenticationError = authenticationError;
            }

            Response.Redirect("~/MainForm.aspx");
        }
    }
}