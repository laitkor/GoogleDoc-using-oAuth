using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI;

namespace Googledocweb
{
    class UtilityCode
    {
        private const string Error_CSS = "color_red";
        private const string Info_CSS = "color_blue";
        private const string Warning_CSS = "color_orange";

        public static void Setmessage(string message, Label lblmsg, MessageType msgtype)
        {
            String table = "<table align='center' cellspacing='2' cellpadding='2'><tr>";
            string img = String.Empty;
            if (msgtype == MessageType.Error)
            {
                lblmsg.CssClass = Error_CSS;
                img = " <img alt='' width='18px' height='18px' class='padding_5' src='" + (HttpContext.Current.Handler as Page).ResolveUrl("~/Image/error.png") + "' />";
            }
            else if (msgtype == MessageType.Information)
            {
                lblmsg.CssClass = Info_CSS;
                img = " <img alt=''  width='18px' height='18px'  class='padding_5' src='" + (HttpContext.Current.Handler as Page).ResolveUrl("~/Image/Information.png") + "' />";
            }
            else
            {
                lblmsg.CssClass = Warning_CSS;
                img = " <img alt=''  width='18px' height='18px'  class='padding_5' src='" + (HttpContext.Current.Handler as Page).ResolveUrl("~/Image/Warning.png") + "' />";
            }
            lblmsg.Text = table + "<td>" + img + "</td>" + "<td>" + message + "</td></tr></table>";

            //lblmsg.Text = "<div class='display_block'>"++"</div>";
        }

        public static string Displaying(int from, int to, int total)
        {
            if (to > total)
                to = total;
            return "Displaying records " + from.ToString() + " - " + to.ToString() + " of " + total.ToString();
        }
        /// <summary>
        /// As user name contains UserName merged with CompanyID alongwith '~' sign.
        /// Using this function only gets UserName...while removing CompanyID from it
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        public static string FormatUserName(string username)
        {
            string[] arr_name = username.Split('~');
            return arr_name[0];
        }
    }
    public enum MessageType
    {
        Information,
        Warning,
        Error
    }
}
