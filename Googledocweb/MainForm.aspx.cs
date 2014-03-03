using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Google.GData.Spreadsheets;
using Google.GData.Client;
using System.Data;
using Google.Documents;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Text;
using System.Web.UI.HtmlControls;
using System.Collections;
using System.Diagnostics;
using System.Reflection;
using Novacode;
using System.Drawing;
using Microsoft.Office.Interop.Owc11;
using System.Xml.Linq;
using DotNetOpenAuth.OpenId.RelyingParty;
using DotNetOpenAuth.OpenId.Extensions.AttributeExchange;
using System.Web.Security;
using System.Security;
using DotNetOpenAuth.OAuth2;
using System.Runtime.InteropServices;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;
using System.Linq.Expressions;
using DotNetOpenAuth.OpenId;
using System.Net;
using Google.GData.Documents;
using DotNetOpenAuth.InfoCard;
namespace Googledocweb
{
    public partial class MainForm : System.Web.UI.Page
    {
        public class GoogleAuthAPICredential
        {
            public string Clientid { get; set; }
            public string Clientkey { get; set; }
            public string Redirecturi { get; set; }

        }
        private GoogleAuthAPICredential local;
        private GoogleAuthAPICredential live;
        string spreadsheetName
        {
            get
            {
                string str = "";
                if (ViewState["spreadsheetName"] != null)
                {
                    str = ViewState["spreadsheetName"].ToString();
                }
                else
                {
                    ViewState["spreadsheetName"] = str;
                }
                return str;
            }
            set
            {
                ViewState["spreadsheetName"] = value;
            }
        }

        string spreadsheetId
        {
            get
            {
                string str = "";
                if (ViewState["spreadsheetId"] != null)
                {
                    str = ViewState["spreadsheetId"].ToString();
                }
                else
                {
                    ViewState["spreadsheetId"] = str;
                }
                return str;
            }
            set
            {
                ViewState["spreadsheetId"] = value;
            }
        }

        string wrkSheetName
        {
            get
            {
                string str = "";
                if (ViewState["wrkSheetName"] != null)
                {
                    str = ViewState["wrkSheetName"].ToString();
                }
                else
                {
                    ViewState["wrkSheetName"] = str;
                }
                return str;
            }
            set
            {
                ViewState["wrkSheetName"] = value;
            }
        }
        string docName
        {
            get
            {
                string str = "";
                if (ViewState["docName"] != null)
                {
                    str = ViewState["docName"].ToString();
                }
                else
                {
                    ViewState["docName"] = str;
                }
                return str;
            }
            set
            {
                ViewState["docName"] = value;
            }
        }
        string uridoc
        {
            get
            {
                string str = "";
                if (ViewState["uridoc"] != null)
                {
                    str = ViewState["uridoc"].ToString();
                }
                else
                {
                    ViewState["uridoc"] = str;
                }
                return str;
            }
            set
            {
                ViewState["uridoc"] = value;
            }
        }
        int docindex
        {
            get
            {
                int str = -1;
                if (ViewState["docindex"] != null)
                {
                    str = (int)ViewState["docindex"];
                }
                else
                {
                    ViewState["docindex"] = str;
                }
                return str;
            }
            set
            {
                ViewState["docindex"] = value;
            }
        }


        string TokenKey
        {
            get
            {
                string str = "";
                if (ViewState["TokenKey"] != null)
                {
                    str = ViewState["TokenKey"].ToString();
                }
                else
                {
                    ViewState["TokenKey"] = str;
                }
                return str;
            }
            set
            {
                ViewState["TokenKey"] = value;
            }
        }
        private bool Resizing = false;
        DataTable myTable;
        static Assembly g_assembly;
        static DocX g_document;
        string email;
        string pass;
        string authorizationUrl;
        string accessToken;
        string refreshtoken;
        int countpage = 0;
        string clientid;
        string clientkey;
        string redirecturi;
        private const string CALLBACK_PARAMETER = "callback";
        private const string RETURNURL_PARAMETER = "ReturnUrl";
        private const string AUTHENTICATION_ENDPOINT = "~/MainForm.aspx";
        private const string GOOGLE_OAUTH_ENDPOINT = "https://www.google.com/accounts/o8/id";
        OAuth2Parameters parameters = new OAuth2Parameters();
        DocumentBuilder builder;
        DocumentBuilder builderefforts;
        DocumentBuilder builderhours;
        DocumentBuilder builderlabor;
        DocumentBuilder builderservice;
        DocumentBuilder builderpayment;
        DocumentBuilder builderImage;
        DocumentBuilder builderArch;
        DocumentsFeed myFeeddoc;
        Aspose.Words.Document doc;
        string[] splittask;
        string[] splitefforts;
        string[] splithours;
        string[] splitlabor;
        string[] splitservice;
        string[] splitpayment;
        string[] splitfields;
        Stream s;
        string SaveLocation;

        protected void Page_Load(object sender, EventArgs e)
        {
            Lbl_Msg.Text = "";
            Lbl_Msg.CssClass = "";
            string host = HttpContext.Current.Request.Url.Host.ToLower();
            if (host == "localhost")
            {
                local = new GoogleAuthAPICredential();
                local.Clientid = "1009066342920-6dea776d0vkq5n4s2q0l7rrdshlpieou.apps.googleusercontent.com";
                local.Clientkey = "a22K0z3n2lV1e05V-eSj_hxX";
                local.Redirecturi = "http://localhost:1705/oauth2callback.aspx";
                clientid = local.Clientid;
                clientkey = local.Clientkey;
                redirecturi = local.Redirecturi;
            }
            else
            {
                live = new GoogleAuthAPICredential();
                live.Clientid = "1009066342920-go89a87imvksg7oq0sor2bbgo1pih941.apps.googleusercontent.com";
                live.Clientkey = "BXbfF6-ls4FqfgARUeEH24ws";
                live.Redirecturi = "http://googledocviewer.dealopia.com/oauth2callback.aspx";
                clientid = live.Clientid;
                clientkey = live.Clientkey;
                redirecturi = live.Redirecturi;
            }
            if (!IsPostBack)
            {
                //if (Session["Token"] != null && !string.IsNullOrWhiteSpace(Session["Token"].ToString()))
                //{
                //    TokenKey = Session["Token"].ToString();
                //    UseToken(TokenKey);
                //}
                if (SessionHelper.Token != null && !string.IsNullOrWhiteSpace(SessionHelper.Token.ToString()))
                {
                    TokenKey = SessionHelper.Token.ToString();
                    UseToken(TokenKey);
                }
                else if (SessionHelper.AuthenticationError != null)
                {
                    if (SessionHelper.AuthenticationError.ToLower() == "access_denied")
                    {
                        Panel_Main.Enabled = false;
                        UtilityCode.Setmessage("Application does not have permission to access Google Doc's. You will be redirected to authentication page.", Lbl_Msg, MessageType.Warning);
                        HtmlMeta meta = new HtmlMeta();
                        meta.HttpEquiv = "Refresh";
                        meta.Content = "3;url=AuthenticationEndPoint.aspx";
                        this.Page.Controls.Add(meta);
                        return;
                    }
                }
                else
                {
                    listView1.Visible = false;
                    Label3.Visible = false;
                    Label1.Visible = false;
                    listView2.Visible = false;
                    button2.Enabled = false;
                    try
                    {
                        SpreadsheetsService GoogleExcelService;
                        GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                        parameters.ClientId = clientid;
                        parameters.ClientSecret = clientkey;
                        string SCOPE = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
                        parameters.Scope = SCOPE;
                        //parameters.RedirectUri = "urn:ietf:wg:oauth:2.0:oob";
                        parameters.RedirectUri = redirecturi;
                        authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
                        Response.Redirect(authorizationUrl);
                    }
                    catch
                    {
                        //ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Invalid Credentials');", true);
                        UtilityCode.Setmessage("Invalid credentials", Lbl_Msg, MessageType.Error);
                        return;
                    }
                }
            }
            else
            {
                hurl.Value = "";
            }
        }

        public void downloadgoogledoc(string Token)
        {
            parameters.ClientId = clientid;
            parameters.ClientSecret = clientkey;
            string SCOPE = "https://docs.google.com/feeds/default/private/full/document";
            parameters.Scope = SCOPE;
            parameters.RedirectUri = redirecturi;
            parameters.AccessCode = Token;
            accessToken = parameters.AccessToken;
            parameters.AccessToken = Token;
            parameters.RefreshToken = Session["Refresh"].ToString();
            //Session["accesstoken"] = accessToken;
            SessionHelper.AccessToken = accessToken;
            GOAuth2RequestFactory requestFactory =
            new GOAuth2RequestFactory(null, "MyDocumentsListIntegration-v1", parameters);
            DocumentsService GoogleExcelService;
            GoogleExcelService = new DocumentsService("Spreadsheet-Abhishek-Test-App");
            GoogleExcelService.RequestFactory = requestFactory;
            DocumentsListQuery query = new DocumentsListQuery();
            myFeeddoc = GoogleExcelService.Query(query);
            DocumentEntry entry = (DocumentEntry)myFeeddoc.Entries[docindex];
            string downloadUrl = entry.Content.Src.Content + "&exportFormat=docx&format=docx";
            bool upl = SaveFileFromURL(downloadUrl, Server.MapPath("Data" + "\\" + myFeeddoc.Entries[docindex].Title.Text + ".docx"));
        }

        public static bool SaveFileFromURL(string url, string destinationFileName)
        {
            // Create a web request to the URL
            HttpWebRequest MyRequest = (HttpWebRequest)WebRequest.Create(url);
            try
            {
                // Get the web response
                HttpWebResponse MyResponse = (HttpWebResponse)MyRequest.GetResponse();

                // Make sure the response is valid
                if (HttpStatusCode.OK == MyResponse.StatusCode)
                {
                    // Open the response stream
                    using (Stream MyResponseStream = MyResponse.GetResponseStream())
                    {
                        // Open the destination file
                        using (FileStream MyFileStream = new FileStream(destinationFileName, FileMode.OpenOrCreate, FileAccess.Write))
                        {
                            // Create a 4K buffer to chunk the file
                            byte[] MyBuffer = new byte[4096];
                            int BytesRead;
                            // Read the chunk of the web response into the buffer

                            while (0 < (BytesRead = MyResponseStream.Read(MyBuffer, 0, MyBuffer.Length)))
                            {
                                // Write the chunk from the buffer to the file
                                MyFileStream.Write(MyBuffer, 0, BytesRead);
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                throw new Exception("Error saving file from URL:" + err.Message, err);
            }
            return true;
        }

        //code to fill the listview of google documents from  google////

        protected void button1_Click(object sender, EventArgs e)
        {
            try
            {

                SpreadsheetsService GoogleExcelService;
                GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                parameters.ClientId = clientid;
                parameters.ClientSecret = clientkey;
                string SCOPE = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
                parameters.Scope = SCOPE;
                parameters.RedirectUri = redirecturi;
                authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            }
            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Invalid Credentials');", true);
                return;
            }


        }

        protected void button2_Click(object sender, EventArgs e)
        {
            if (spreadsheetName.Length > 0)
            {
                if (listView2.SelectedIndex > -1)
                {
                    wrkSheetName = listView2.SelectedItem.Text;
                }
                else
                {
                    // return;
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select Worksheet First');", true);
                    return;
                }
                parameters.ClientId = clientid;
                parameters.ClientSecret = clientkey;
                string SCOPE = "http://spreadsheets.google.com/feeds https://docs.google.com/feeds";
                parameters.Scope = SCOPE;
                parameters.RedirectUri = redirecturi;
                // parameters.RedirectUri = "https://googledocviewer.dealopia.com/oauth2callback";
                parameters.AccessToken = SessionHelper.AccessToken; //Session["accesstoken"].ToString();
                SpreadsheetsService GoogleExcelService;
                GOAuth2RequestFactory requestFactory =
                new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
                GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                GoogleExcelService.RequestFactory = requestFactory;
                Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();
                SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
                if (myFeed == null)
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Invalid Credentials');", true);
                }

                foreach (SpreadsheetEntry mySpread in myFeed.Entries)
                {
                    if (mySpread.Title.Text == spreadsheetName)
                    {
                        WorksheetFeed wfeed = mySpread.Worksheets;
                        foreach (WorksheetEntry wsheet in wfeed.Entries)
                        {
                            if (wsheet.Title.Text == wrkSheetName)
                            {
                                AtomLink atm = wsheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);
                                ListQuery Lquery = new ListQuery(atm.HRef.ToString());
                                ListFeed LFeed = GoogleExcelService.Query(Lquery);
                                myTable = new DataTable();
                                DataColumn DC;
                                foreach (ListEntry LmySpread in LFeed.Entries)
                                {
                                    DataRow myDR = myTable.NewRow();
                                    foreach (ListEntry.Custom listrow in LmySpread.Elements)
                                    {
                                        DC = myTable.Columns[listrow.LocalName] ?? myTable.Columns.Add(listrow.LocalName);
                                        myDR[DC] = listrow.Value;
                                    }
                                    myTable.Rows.Add(myDR);
                                }
                                dataGridView1.DataSource = myTable;
                                dataGridView1.DataBind();
                                ExportToPdf(myTable);
                            }
                        }
                    }
                }
            }
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Data Reading is Completed');", true);
        }

        public void bindgrid(string Token)
        {
            parameters.ClientId = clientid;
            parameters.ClientSecret = clientkey;
            string SCOPE = "http://spreadsheets.google.com/feeds https://docs.google.com/feeds";
            parameters.Scope = SCOPE;
            parameters.RedirectUri = redirecturi;
            //parameters.RedirectUri = "https://googledocviewer.dealopia.com/oauth2callback";
            parameters.AccessToken = SessionHelper.AccessToken;//Session["accesstoken"].ToString();
            GOAuth2RequestFactory requestFactory = new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
            SpreadsheetsService GoogleExcelService;
            GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
            GoogleExcelService.RequestFactory = requestFactory;
            Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();
            SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
            foreach (SpreadsheetEntry mySpread in myFeed.Entries)
            {
                //if (mySpread.Title.Text == spreadsheetName)
                //{
                if (mySpread.Id.AbsoluteUri == spreadsheetId)
                {
                    WorksheetFeed wfeed = mySpread.Worksheets;
                    foreach (WorksheetEntry wsheet in wfeed.Entries)
                    {
                        if (wsheet.Title.Text == wrkSheetName)
                        {
                            AtomLink atm = wsheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);
                            ListQuery Lquery = new ListQuery(atm.HRef.ToString());
                            ListFeed LFeed = GoogleExcelService.Query(Lquery);
                            myTable = new DataTable();
                            DataColumn DC;
                            foreach (ListEntry LmySpread in LFeed.Entries)
                            {
                                DataRow myDR = myTable.NewRow();
                                foreach (ListEntry.Custom listrow in LmySpread.Elements)
                                {
                                    DC = myTable.Columns[listrow.LocalName] ?? myTable.Columns.Add(listrow.LocalName);
                                    myDR[DC] = listrow.Value;
                                }
                                myTable.Rows.Add(myDR);
                            }
                            dataGridView1.DataSource = myTable;
                            dataGridView1.DataBind();
                        }
                    }
                }
            }
        }

        public void ExportToPdf(DataTable myDataTable)
        {
            bindgrid(TokenKey);
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 10, 10, 10, 10);
            try
            {
                PdfWriter.GetInstance(pdfDoc, System.Web.HttpContext.Current.Response.OutputStream);
                pdfDoc.Open();
                Chunk c = new Chunk(" Spreadsheet Sample ", FontFactory.GetFont("Verdana", 11));
                iTextSharp.text.Paragraph p = new iTextSharp.text.Paragraph();
                p.Alignment = Element.ALIGN_CENTER;
                p.Add(c);
                pdfDoc.Add(p);
                Chunk cblank = new Chunk(" ", FontFactory.GetFont("Verdana", 11));
                iTextSharp.text.Paragraph pblank = new iTextSharp.text.Paragraph();
                pblank.Alignment = Element.ALIGN_CENTER;
                pblank.Add(cblank);
                pdfDoc.Add(pblank);
                iTextSharp.text.Font font8 = FontFactory.GetFont("ARIAL", 7);
                DataTable dt = myDataTable;
                if (dt != null)
                {
                    //Craete instance of the pdf table and set the number of column in that table  
                    PdfPTable PdfTable = new PdfPTable(dt.Columns.Count);
                    PdfPCell PdfPCell = null;
                    for (int rows = 0; rows < dt.Rows.Count; rows++)
                    {
                        for (int column = 0; column < dt.Columns.Count; column++)
                        {
                            PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Rows[rows][column].ToString(), font8)));
                            PdfTable.AddCell(PdfPCell);
                        }
                    }
                    //PdfTable.SpacingBefore = 15f; // Give some space after the text or it may overlap the table            
                    pdfDoc.Add(PdfTable); // add pdf table to the document   
                }
                pdfDoc.Close();
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "attachment; filename= Spreadsheet.pdf");
                System.Web.HttpContext.Current.Response.Write(pdfDoc);
                Response.Flush();
                Response.End();
                //HttpContext.Current.ApplicationInstance.CompleteRequest();  
            }
            catch (DocumentException de)
            {
                System.Web.HttpContext.Current.Response.Write(de.Message);
            }
            catch (IOException ioEx)
            {
                System.Web.HttpContext.Current.Response.Write(ioEx.Message);
            }
            catch (Exception ex)
            {
                System.Web.HttpContext.Current.Response.Write(ex.Message);
            }
        }

        public void createpdf(DataTable myTable)
        {
            bindgrid(TokenKey);
            /*Generate Filename in serial order*/
            DateTime dtNow = DateTime.Now;
            string FName = "Invoice Requisition";
            iTextSharp.text.Document doc = new iTextSharp.text.Document();
            PdfWriter.GetInstance(doc, System.Web.HttpContext.Current.Response.OutputStream);
            var document = new iTextSharp.text.Document();
            string path = Server.MapPath("~/Invoice");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var fileName = "Invoice.Pdf";
            var fullPath = path + "\\" + DateTime.Now.ToString("HH.mm.ss") + fileName;
            FileInfo fi = new FileInfo(fullPath);
            if (fi.Exists)
            {
                fi.Delete();
            }
            if (File.Exists(fullPath))
            {
                fileName = DateTime.Now.ToString("HH.mm.ss") + fileName;
                fullPath = path + "\\" + fileName;

            }
            else
            {

            }
            PdfWriter.GetInstance(doc, new FileStream(fullPath, FileMode.Create));
            doc.Open();
            iTextSharp.text.Font font12 = FontFactory.GetFont("ARIAL", 12);
            iTextSharp.text.Font font9 = FontFactory.GetFont("ARIAL", 9);
            iTextSharp.text.Font font8 = FontFactory.GetFont("ARIAL", 8);
            iTextSharp.text.Font font7 = FontFactory.GetFont("ARIAL", 7);

            /*Create instance of the pdf table and set the number of column in that table*/

            PdfPTable PdfTable = new PdfPTable(4);
            PdfPCell PdfPCell = null;
            string str = "";
            #region Invoice heading
            PdfPCell PdfPCellBlank = new PdfPCell(new Phrase(new Chunk("  ", font9)));
            PdfPCellBlank.Colspan = 4;
            PdfPCellBlank.Border = 0;
            PdfTable.AddCell(PdfPCellBlank);
            PdfPCell PdfPCellInvoice = new PdfPCell(new Phrase(new Chunk("Company Employee Details", font12)));
            PdfPCellInvoice.Colspan = 4;
            PdfPCellInvoice.Border = 0;
            PdfPCellInvoice.HorizontalAlignment = Element.ALIGN_CENTER;
            PdfPCellInvoice.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            PdfTable.AddCell(PdfPCellInvoice);
            PdfTable.AddCell(PdfPCellBlank);
            #endregion Invoice heading
            #region companyinfo

            PdfPCell PdfPCellPrac = new PdfPCell(new Phrase(new Chunk("Company Information", font9)));
            PdfPCellPrac.Colspan = 4;
            PdfPCellPrac.Border = 0;
            PdfPCellPrac.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            PdfTable.AddCell(PdfPCellPrac);
            str = "Name: " + myTable.Rows[0]["Name"].ToString();
            PdfPCellPrac = new PdfPCell(new Phrase(new Chunk(str, font7)));
            PdfPCellPrac.Colspan = 4;
            PdfPCellPrac.Border = 0;
            PdfTable.AddCell(PdfPCellPrac);
            str = "Address: " + myTable.Rows[0]["Address"].ToString();
            PdfPCellPrac = new PdfPCell(new Phrase(new Chunk(str, font7)));
            PdfPCellPrac.Colspan = 4;
            PdfPCellPrac.Border = 0;
            PdfTable.AddCell(PdfPCellPrac);
            str = "City: " + myTable.Rows[0]["City"].ToString();
            PdfPCellPrac = new PdfPCell(new Phrase(new Chunk(str, font7)));
            PdfPCellPrac.Colspan = 4;
            PdfPCellPrac.Border = 0;
            PdfTable.AddCell(PdfPCellPrac);
            #endregion

            #region projectinfo

            PdfTable.AddCell(PdfPCellBlank);
            PdfPCell PdfPCellProject = new PdfPCell(new Phrase(new Chunk("Contact Detail", font9)));
            PdfPCellProject.Colspan = 4;
            PdfPCellProject.Border = 0;
            PdfPCellProject.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            PdfTable.AddCell(PdfPCellProject);
            str = "Phone: " + myTable.Rows[0]["Phone"].ToString();
            PdfPCellProject = new PdfPCell(new Phrase(new Chunk(str, font7)));
            PdfPCellProject.Colspan = 4;
            PdfPCellProject.Border = 0;
            PdfTable.AddCell(PdfPCellProject);
            PdfTable.AddCell(PdfPCellBlank);

            #endregion

            #region DeveloperInfo

            PdfTable.AddCell(PdfPCellBlank);
            PdfPCell PdfPCellDeveloper = new PdfPCell(new Phrase(new Chunk("Personal Details", font9)));
            PdfPCellDeveloper.Colspan = 4;
            PdfPCellDeveloper.Border = 0;
            PdfPCellDeveloper.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            PdfTable.AddCell(PdfPCellDeveloper);

            /* Add header to Developer table*/

            PdfTable.AddCell(PdfPCellBlank);
            PdfPCell = new PdfPCell(new Phrase(new Chunk("Name", font8)));
            PdfPCell.Colspan = 1;
            PdfTable.AddCell(PdfPCell);
            PdfPCell = new PdfPCell(new Phrase(new Chunk("City", font8)));
            PdfPCell.Colspan = 1;
            PdfTable.AddCell(PdfPCell);
            PdfPCell = new PdfPCell(new Phrase(new Chunk("Address", font8)));
            PdfPCell.Colspan = 1;
            PdfTable.AddCell(PdfPCell);
            PdfPCell = new PdfPCell(new Phrase(new Chunk("Phone", font8)));
            PdfPCell.Colspan = 1;
            PdfTable.AddCell(PdfPCell);


            /* Add Developer Name From dataset*/

            for (int rows = 0; rows < myTable.Rows.Count; rows++)
            {
                string devname = myTable.Rows[rows]["Name"].ToString();
                PdfPCell = new PdfPCell(new Phrase(new Chunk(devname, font7)));
                PdfPCell.Colspan = 1;
                PdfTable.AddCell(PdfPCell);
                string date = myTable.Rows[rows]["City"].ToString();
                PdfPCell = new PdfPCell(new Phrase(new Chunk(date, font7)));
                PdfPCell.Colspan = 1;
                PdfTable.AddCell(PdfPCell);
                string address = myTable.Rows[rows]["Address"].ToString();
                PdfPCell = new PdfPCell(new Phrase(new Chunk(address, font7)));
                PdfPCell.Colspan = 1;
                PdfTable.AddCell(PdfPCell);
                string workinghrs = myTable.Rows[rows]["Phone"].ToString();
                PdfPCell = new PdfPCell(new Phrase(new Chunk(workinghrs, font7)));
                PdfPCell.Colspan = 1;
                PdfTable.AddCell(PdfPCell);
            }
            PdfTable.AddCell(PdfPCellBlank);
            #endregion
            #region
            /* Displaying Total Cost and total working Hours*/
            #endregion
            #region
            PdfPCell PdfPCellComment = new PdfPCell(new Phrase(new Chunk("Comment", font9)));
            PdfPCellComment.Colspan = 4;
            PdfPCellComment.Border = 0;
            PdfPCellComment.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            PdfTable.AddCell(PdfPCellComment);
            //str = Txtcmmt.Text;
            PdfPCellProject = new PdfPCell(new Phrase(new Chunk(str, font7)));
            PdfPCellProject.Colspan = 4;
            PdfPCellProject.Border = 0;
            PdfTable.AddCell(PdfPCellProject);
            #endregion
            #region sign
            #endregion sign
            doc.Add(PdfTable);
            #region footer
            Chunk myFooter = new Chunk(DateTime.Now.ToShortDateString() + "      " + "Page " + (doc.PageNumber + 1),
               FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8, new iTextSharp.text.Color(46, 84, 141)));
            iTextSharp.text.HeaderFooter footer = new iTextSharp.text.HeaderFooter(new Phrase(myFooter), false);    //Create a footer object with the chunk data
            footer.Border = iTextSharp.text.Rectangle.NO_BORDER;    //Specify no border around the footer
            footer.Alignment = Element.ALIGN_RIGHT; //Specify the footer text alignment
            doc.Footer = footer;                  //Set the document's footer to the footer object

            #endregion
            doc.Close();

            Response.ContentType = "Application/pdf";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
            String assemblyPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            Process proc = new Process();
            proc.StartInfo = new ProcessStartInfo()
            {
                FileName = fullPath
            };

            Response.Flush();
            Response.End();

        }

        public void CreateWordTableWithDataTable(DataTable myTable)
        {

            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];
                    } //end row loop
                } //end column loop

                Microsoft.Office.Interop.Word.Document Doc = new Microsoft.Office.Interop.Word.Document();
                Doc.Application.Visible = true;
                Doc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;
                dynamic Range = Doc.Content.Application.Selection.Range;
                String Temp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        Temp = Temp + DataArray[r, c] + "\t";

                    }
                }
                Range.Text = Temp;
                object Separator = Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
                object Format = Microsoft.Office.Interop.Word.WdTableFormat.wdTableFormatWeb1;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent;
                Range.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount, Type.Missing, ref Format, ref ApplyBorders, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, ref AutoFit, ref AutoFitBehavior,
                Type.Missing);
                Range.Select();
                Doc.Application.Selection.Tables[1].Select();
                Doc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                Doc.Application.Selection.Tables[1].Rows.Alignment = 0;
                Doc.Application.Selection.Tables[1].Rows[1].Select();
                Doc.Application.Selection.InsertRowsAbove(1);
                Doc.Application.Selection.Tables[1].Rows[1].Select();
                //gotta do the header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    Doc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = myTable.Columns[c].ColumnName;
                }
                Doc.Application.Selection.Tables[1].Rows[1].Select();
                Doc.Application.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
            catch (Exception ex)
            {
                // throw ex;
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('No Data To be Exported');", true);
            }
        }

        protected void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                if (listView1.SelectedIndex > -1)
                {
                    spreadsheetName = listView1.SelectedItem.Text;
                    spreadsheetId = listView1.SelectedItem.Value;
                }
                else
                    // return;
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select Spreadsheet First');", true);
                if (spreadsheetName.Length > 0)
                {
                    //SpreadsheetsService GoogleExcelService;
                    //GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                    //GoogleExcelService.setUserCredentials(Session["Email"].ToString(), Session["Pass"].ToString());
                    parameters.ClientId = clientid;
                    parameters.ClientSecret = clientkey;
                    string SCOPE = "http://spreadsheets.google.com/feeds https://docs.google.com/feeds";
                    parameters.Scope = SCOPE;
                    parameters.RedirectUri = redirecturi;
                    // parameters.RedirectUri = "https://googledocviewer.dealopia.com/oauth2callback";
                    parameters.AccessToken = SessionHelper.AccessToken; //Session["accesstoken"].ToString();
                    GOAuth2RequestFactory requestFactory =
                    new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
                    SpreadsheetsService GoogleExcelService;
                    GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                    //GoogleExcelService.setUserCredentials(Session["Email"].ToString(), Session["Pass"].ToString()); 
                    GoogleExcelService.RequestFactory = requestFactory;
                    Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();
                    SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
                    Label1.Visible = true;
                    listView2.Visible = true;
                    foreach (SpreadsheetEntry mySpread in myFeed.Entries)
                    {
                        //if (mySpread.Title.Text == spreadsheetName)
                        //{
                        if (mySpread.Id.AbsoluteUri == spreadsheetId)
                        {
                            WorksheetFeed wfeed = mySpread.Worksheets;
                            listView2.Items.Clear();
                            foreach (WorksheetEntry wsheet in wfeed.Entries)
                            {
                                string[] row = { wsheet.Title.Text, wsheet.Cols.ToString(), wsheet.Rows.ToString(), wsheet.Summary.Text };
                                System.Web.UI.WebControls.ListItem listItem = new System.Web.UI.WebControls.ListItem(wsheet.Title.Text);
                                listView2.Items.Add(listItem);
                            }
                        }

                    }

                }
            }
            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Invalid Credentials');", true);
                return;
            }

        }

        protected void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            wrkSheetName = "";
            if (listView2.SelectedIndex > -1)
            {
                wrkSheetName = listView2.SelectedItem.Text;

            }
            else
            {
                // return;
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select Worksheet First');", true);
            }
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            //Session["Email"] = "";
            //Session["Pass"] = "";
            ////txtemail.Text = "";
            ////txtPass.Text = "";
            //listView1.Visible = false;
            //listView2.Visible = false;
            //Label1.Visible = false;
            //Label3.Visible = false;
            //ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('User Log Out Successfully');", true);
            //return;

        }

        protected void TextBox1_TextChanged(object sender, EventArgs e)
        {
            UseToken(TokenKey);
        }

        private void UseToken(string Token)
        {
            try
            {
                Session.Clear();
                parameters.ClientId = clientid;
                parameters.ClientSecret = clientkey;
                string SCOPE = "http://spreadsheets.google.com/feeds https://docs.google.com/feeds";
                parameters.Scope = SCOPE;
                parameters.RedirectUri = redirecturi;
                parameters.AccessCode = Token;
                OAuthUtil.GetAccessToken(parameters);
                accessToken = parameters.AccessToken;
                // parameters.AccessToken = Session["Token"].ToString();
                //Session["accesstoken"] = accessToken;
                SessionHelper.AccessToken = accessToken;
                GOAuth2RequestFactory requestFactory =
                new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
                SpreadsheetsService GoogleExcelService;
                GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                GoogleExcelService.RequestFactory = requestFactory;
                Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();
                SpreadsheetFeed myFeed = GoogleExcelService.Query(query);


                foreach (SpreadsheetEntry mySpread in myFeed.Entries)
                {
                    listView1.Visible = true;
                    Label3.Visible = true;
                    button2.Enabled = true;
                    System.Web.UI.WebControls.ListItem listItem = new System.Web.UI.WebControls.ListItem(mySpread.Title.Text, mySpread.Id.AbsoluteUri);
                    listView1.Items.Add(listItem);
                }
            }
            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Time out');", true);
            }
        }

        protected void Button3_Click1(object sender, EventArgs e)
        {
            //SpreadsheetsService GoogleExcelService;
            //GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
            //// GoogleExcelService.setUserCredentials(Session["Email"].ToString(), Session["Pass"].ToString());   
            //// OAuth2Parameters parameters = new OAuth2Parameters();
            //parameters.ClientId = "1009066342920-7i9i3tr89h37gotsu4vst0dk7aun58bh.apps.googleusercontent.com";
            //parameters.ClientSecret = "jQesTK8iLDoM0ItWUpDH54EX";
            //string SCOPE = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
            //parameters.Scope = SCOPE;
            //parameters.RedirectUri = "urn:ietf:wg:oauth:2.0:oob";
            //// parameters.RedirectUri = "";
            //authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            ////TextBox2.Text = authorizationUrl;
            //hurl.Value = authorizationUrl;
            //Response.Write("<script type='text/javascript'>window.open('"+authorizationUrl+"');</script>");
        }

        protected void Button3_Click2(object sender, EventArgs e)
        {


        }

        protected void Button3_Click3(object sender, EventArgs e)
        {


        }

        protected void Button4_Click1(object sender, EventArgs e)
        {
            if (listView1.SelectedIndex > -1 && listView2.SelectedIndex > -1)
            {
                templatecreation(myTable);
            }
            else
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the Worksheet');", true);
            }

        }

        public void templatecreation(DataTable mytable)
        {

            try
            {
                bindgrid(TokenKey);

                //the process cannot access the file because it is being used c#.//
                GC.Collect();
                GC.WaitForPendingFinalizers(); //////  

                if (!ValidateSpreadSheet())
                {
                    UtilityCode.Setmessage("Selected spreadsheet doesnt seems to have valid format. Please select appropriate spreadsheet to create proposal", Lbl_Msg, MessageType.Warning);
                    return;
                }

                splittask = myTable.Rows[FindRow("TaskRange")]["Value"].ToString().Split(',');
                splitefforts = myTable.Rows[FindRow("EffortsRange")]["Value"].ToString().Split(',');
                splithours = myTable.Rows[FindRow("HoursRange")]["Value"].ToString().Split(',');
                splitservice = myTable.Rows[FindRow("ServiceRange")]["Value"].ToString().Split(',');
                splitlabor = myTable.Rows[FindRow("LabourRange")]["Value"].ToString().Split(',');
                splitpayment = myTable.Rows[FindRow("PaymentRange")]["Value"].ToString().Split(',');
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                License li = new License();
                li.SetLicense("Aspose.Words.lic");
                Object Missing = System.Reflection.Missing.Value;
                Object True = true;
                Object False = false;
                string MyDir = Server.MapPath("ProjectPrposal.docx");
                string destination = Server.MapPath("~/NewTemplate/") + "NewProjectProposal.docx";
                doc = new Aspose.Words.Document(MyDir);
                //code for inserting estimated efforts in word document in tabular form////
                InsertLogo();
                InsertTask();
                InsertArchImage();
                InsertEfforts();
                InsertHoursRate();
                InsertLabour();
                Insertservice();
                Insertpayment();

                doc.MailMerge.Execute(
                    new string[] 
                    { 
                        "ProposalHeadline",
                        "TechnicalPartByName",
                        "FinancialPartByName",
                        "ProposalDate",
                        "Title", 
                        "Description", 
                        "Technology", 
                        "Company", 
                        "Address", 
                        "Founded", 
                        "ArchText", 
                        "Technical", 
                        "Requirements", 
                        "Risk", 
                        "Documentation", 
                        "Communication", 
                        "KeyService", 
                        "Skills", 
                        "Location", 
                        "Email", 
                        "Paymentnote", 
                        "Warranty", 
                        "Maintenance", 
                        "Employees",
                        "MY_COMPANY",
                        "MY_COMPANY_ADDRESS"
                    },
                    new object[] 
                    {
                        myTable.Rows[FindRow("ProposalHeadLine")]["Value"].ToString(), 
                        myTable.Rows[FindRow("TechnicalPartByName")]["Value"].ToString(), 
                        myTable.Rows[FindRow("FinancialPartByName")]["Value"].ToString(), 
                        myTable.Rows[FindRow("ProposalDate")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Title")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Project_Description")]["Value"].ToString(),
                        myTable.Rows[FindRow("Technology")]["Value"].ToString(), 
                        myTable.Rows[FindRow("My_Company")]["Value"].ToString(), 
                        myTable.Rows[FindRow("My_Company_Address")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Founded")]["Value"].ToString(), 
                        myTable.Rows[FindRow("ArchText")]["Value"].ToString(),
                        myTable.Rows[FindRow("Technical")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Requirements")]["Value"].ToString(),  
                        myTable.Rows[FindRow("Risk")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Documentation")]["Value"].ToString(), 
                        myTable.Rows[FindRow("Communication")]["Value"].ToString(),
                        myTable.Rows[FindRow("KeyService")]["Value"].ToString(),
                        myTable.Rows[FindRow("KeySkills")]["Value"].ToString(),
                        myTable.Rows[FindRow("Location")]["Value"].ToString(),
                        myTable.Rows[FindRow("Email")]["Value"].ToString(),
                        myTable.Rows[FindRow("Paymentnote")]["Value"].ToString(),
                        myTable.Rows[FindRow("Warranty")]["Value"].ToString(),
                        myTable.Rows[FindRow("Maintenance")]["Value"].ToString(),
                        myTable.Rows[FindRow("Employees")]["Value"].ToString(),
                        myTable.Rows[FindRow("My_Company")]["Value"].ToString(),
                        myTable.Rows[FindRow("My_Company_Address")]["Value"].ToString(),
                    });

                CreateHeaderAndFooter();

                doc.Save(destination);
                Response.ContentType = "application/octetstream";
                Response.AppendHeader("Content-Disposition", "attachment; filename=NewProjectProposal.docx");
                Response.TransmitFile(destination);
                Response.End();
            }
            catch (Exception Ex)
            {
                //ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the Spreadsheet and Template data worksheet');", true);
                UtilityCode.Setmessage("An error occured while creating proposal. Exception - " + Ex.Message, Lbl_Msg, MessageType.Warning);
            }
        }

        private int FindRow(string field)
        {
            int index = -1;
            for (int i = 0; i < myTable.Rows.Count; i++)
            {
                if (myTable.Rows[i]["field"].ToString().Trim() == field)
                {
                    index = i;
                    break;
                }
            }

            return index;
        }

        public void InsertTask()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builder = new DocumentBuilder(doc);
                builder.MoveToMergeField("Task");
                // setting format of builder of task//
                formattable(builder);
                ///
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splittask[0]);
                //ColumnCount - 1
                for (int c = 0; c <= 3; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splittask[0]) - 2; r <= Convert.ToInt32(splittask[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];

                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splittask[0]) - 2; r <= Convert.ToInt32(splittask[1]) - 2; r++)
                {
                    string str = DataArray[r, 0].ToString();
                    for (int i = 0; i < str.Length; i++)
                    {
                        int asciival = ((int)str[i]);
                        if (asciival >= 48 && asciival <= 57)
                        {
                            builder.Font.Bold = false;
                            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                            builder.RowFormat.Alignment = RowAlignment.Center;


                            // builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                        }
                        else
                        {
                            if (r == Convert.ToInt32(splittask[0]) - 2)
                            {
                                builder.Font.Bold = true;
                                builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                                builder.CellFormat.Width = 5;
                                builder.RowFormat.Height = 5;
                                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                                builder.RowFormat.Alignment = RowAlignment.Center;
                            }
                            else
                            {
                                builder.Font.Bold = true;
                                builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                                builder.CellFormat.Width = 5;
                                builder.RowFormat.Height = 5;
                                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                                builder.RowFormat.Alignment = RowAlignment.Center;

                            }
                        }
                    }
                    for (int c = 0; c <= 3; c++)
                    {
                        builder.InsertCell();

                        //making the text center align////////
                        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builder.RowFormat.Alignment = RowAlignment.Center;

                        //formatting the width of hrs column in Task table///
                        if (c == 0 || c == 2)
                        {
                            builder.CellFormat.Width = 20.0;
                            builder.RowFormat.Height = 5;
                            // code for making text inside cell center align/////
                            // builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        }



                        else
                        {
                            builder.CellFormat.Width = 200.0;
                            builder.RowFormat.Height = 5;


                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builder.Write(Temp);
                        Temp = "";
                    }
                    builder.EndRow();
                }
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }
        }

        //public void InsertEfforts()
        //{
        //    try
        //    {
        //        int RowCount = myTable.Rows.Count;
        //        int ColumnCount = myTable.Columns.Count;
        //        builderefforts = new DocumentBuilder(doc);
        //        builderefforts.MoveToMergeField("Efforts");
        //        //format of table of efforts///
        //        formattable(builderefforts);
        //        Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
        //        int r = Convert.ToInt32(splitefforts[0]);
        //        for (int c = 0; c <= 3; c++)
        //        {
        //            DataArray[r, c] = myTable.Columns[c].ColumnName;
        //            for (r = Convert.ToInt32(splitefforts[0]) - 2; r <= Convert.ToInt32(splitefforts[1]) - 2; r++)
        //            {
        //                DataArray[r, c] = myTable.Rows[r][c];

        //            }
        //        }
        //        String Temp = "";
        //        for (r = Convert.ToInt32(splitefforts[0]) - 2; r <= Convert.ToInt32(splitefforts[1]) - 2; r++)
        //        {
        //            if (r % 2 == 0)
        //            {
        //                builderefforts.Font.Bold = false;
        //                builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
        //            }
        //            else
        //            {
        //                //this is for heading placed at the estimated efforts
        //                if (r == Convert.ToInt32(splitefforts[0]) - 2)
        //                {
        //                    builderefforts.Font.Bold = true;
        //                    builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
        //                }
        //                else
        //                {
        //                    builderefforts.Font.Bold = false;
        //                    builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
        //                    //builder.CellFormat.Width = 5;
        //                    //builder.RowFormat.Height = 5;
        //                    //builderefforts.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        //                    //builderefforts.RowFormat.Alignment = RowAlignment.Center; 
        //                }
        //            }

        //            for (int c = 0; c <= 3; c++)
        //            {
        //                builderefforts.InsertCell();
        //                //builderefforts.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        //                //builderefforts.RowFormat.Alignment = RowAlignment.Center; 
        //                //formatting the width of the hrs column///
        //                if (c == 0 || c == 2)
        //                {
        //                    builderefforts.CellFormat.Width = 20.0;
        //                }
        //                else
        //                {
        //                    builderefforts.CellFormat.Width = 200.0;
        //                }
        //                Temp = Temp + DataArray[r, c] + "\t";
        //                builderefforts.Write(Temp);
        //                Temp = "";
        //            }
        //            builderefforts.EndRow();
        //        }
        //    }
        //    catch
        //    {
        //        ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
        //    }
        //}

        public void InsertEfforts()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderefforts = new DocumentBuilder(doc);
                builderefforts.MoveToMergeField("Efforts");

                ////format of table of efforts///
                formattable(builderefforts);

                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splitefforts[0]);
                for (int c = 0; c <= 3; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splitefforts[0]) - 2; r <= Convert.ToInt32(splitefforts[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];
                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splitefforts[0]) - 2; r <= Convert.ToInt32(splitefforts[1]) - 2; r++)
                {
                    if (r % 2 == 0)
                    {
                        builderefforts.Font.Bold = false;
                        builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                    }
                    else
                    {
                        //this is for heading placed at the estimated efforts
                        if (r != Convert.ToInt32(splitefforts[0]) - 2)
                        {
                            builderefforts.Font.Bold = false;
                            builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                        }
                    }

                    for (int c = 0; c <= 3; c++)
                    {
                        builderefforts.InsertCell();

                        builderefforts.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builderefforts.RowFormat.Alignment = RowAlignment.Center;

                        if (r == Convert.ToInt32(splitefforts[0]) - 2)
                        {
                            builderefforts.Font.Bold = true;
                            builderefforts.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                        }

                        if (c == 0 || c == 2)
                        {
                            builderefforts.CellFormat.Width = 50.0;
                            builder.RowFormat.Height = 5;
                        }
                        else
                        {
                            builderefforts.CellFormat.Width = 200.0;
                            builder.RowFormat.Height = 5;
                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builderefforts.Write(Temp);
                        Temp = "";
                    }
                    builderefforts.EndRow();
                }
                builderefforts.EndTable();
            }
            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }
        }

        public void InsertLogo()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderImage = new DocumentBuilder(doc);
                builderImage.MoveToMergeField("Logo");
                builderImage.StartTable();
                builderImage.RowFormat.Height = 20;
                builderImage.RowFormat.HeightRule = HeightRule.AtLeast;
                // Some special features for the header row.
                // builderImage.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.FromArgb(198, 217, 241);
                // other alignment setting
                builderImage.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builderImage.CellFormat.Width = 80;
                builderImage.RowFormat.Height = 80;
                builderImage.RowFormat.HeightRule = HeightRule.Auto;
                // making the image align center//////
                builderImage.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                builderImage.RowFormat.Alignment = RowAlignment.Center;
                builderImage.InsertCell();
                var html = "<div><img src='" + myTable.Rows[FindRow("Logo")]["Value"].ToString() + "'/></div>";
                builderImage.InsertHtml(html);
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }

        }

        public void InsertArchImage()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderArch = new DocumentBuilder(doc);
                builderArch.MoveToMergeField("SystemArchImage");
                builderArch.StartTable();
                builderArch.RowFormat.Height = 20;
                builderArch.RowFormat.HeightRule = HeightRule.AtLeast;
                // Some special features for the header row.
                // builderImage.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.FromArgb(198, 217, 241);
                // other alignment setting
                builderArch.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builderArch.CellFormat.Width = 200;
                builderArch.RowFormat.Height = 200;
                builderArch.RowFormat.HeightRule = HeightRule.Auto;
                // making the image align center//////
                builderArch.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                builderArch.RowFormat.Alignment = RowAlignment.Center;
                builderArch.InsertCell();
                var html = @"<div><img src='" + myTable.Rows[FindRow("ArchitectureImage")]["Value"].ToString() + "'/></div>";
                builderArch.InsertHtml(html);
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }
        }

        public void InsertHoursRate()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderhours = new DocumentBuilder(doc);
                builderhours.MoveToMergeField("HoursRate");
                formattable(builderhours);
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splithours[0]);
                for (int c = 0; c <= 3; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splithours[0]) - 2; r <= Convert.ToInt32(splithours[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];
                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splithours[0]) - 2; r <= Convert.ToInt32(splithours[1]) - 2; r++)
                {
                    if (r % 2 == 0)
                    {
                        builderhours.Font.Bold = false;
                        builderhours.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                    }
                    else
                    {
                        //this is for heading placed at the estimated efforts
                        if (r != Convert.ToInt32(splithours[0]) - 2)
                        {
                            builderhours.Font.Bold = false;
                            builderhours.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                        }
                    }

                    for (int c = 0; c <= 3; c++)
                    {
                        builderhours.InsertCell();
                        builderhours.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builderhours.RowFormat.Alignment = RowAlignment.Center;

                        if (r == Convert.ToInt32(splithours[0]) - 2)
                        {
                            builderhours.Font.Bold = true;
                            builderhours.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                        }

                        if (c == 0 || c == 2)
                        {
                            builderhours.CellFormat.Width = 50.0;
                        }
                        else
                        {
                            builderhours.CellFormat.Width = 200.0;
                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builderhours.Write(Temp);
                        Temp = "";
                    }
                    builderhours.EndRow();
                }
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }

        }

        public void InsertLabour()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderlabor = new DocumentBuilder(doc);
                builderlabor.MoveToMergeField("Labor");
                formattable(builderlabor);
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splitlabor[0]);
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splitlabor[0]) - 2; r <= Convert.ToInt32(splitlabor[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];

                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splitlabor[0]) - 2; r <= Convert.ToInt32(splitlabor[1]) - 2; r++)
                {
                    if (r % 2 == 0)
                    {
                        builderlabor.Font.Bold = false;
                        builderlabor.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                    }
                    else
                    {
                        //this is for heading placed at the estimated efforts
                        if (r != Convert.ToInt32(splitlabor[0]) - 2)
                        {
                            builderlabor.Font.Bold = false;
                            builderlabor.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                        }
                    }

                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        builderlabor.InsertCell();
                        builderlabor.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builderlabor.RowFormat.Alignment = RowAlignment.Center;

                        if (r == Convert.ToInt32(splitlabor[0]) - 2)
                        {
                            builderlabor.Font.Bold = true;
                            builderlabor.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                        }

                        if (c == 0)
                        {
                            builderlabor.CellFormat.Width = 50.0;
                        }
                        else
                        {
                            builderlabor.CellFormat.Width = 200.0;
                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builderlabor.Write(Temp);
                        Temp = "";
                    }
                    builderlabor.EndRow();
                }
            }

            catch
            {

                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }

        }

        public void Insertservice()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderservice = new DocumentBuilder(doc);
                builderservice.MoveToMergeField("Service");
                //format of table of efforts///
                formattable(builderservice);
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splitservice[0]);
                for (int c = 0; c <= 3; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splitservice[0]) - 2; r <= Convert.ToInt32(splitservice[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];

                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splitservice[0]) - 2; r <= Convert.ToInt32(splitservice[1]) - 2; r++)
                {
                    if (r % 2 == 0)
                    {
                        builderservice.Font.Bold = false;
                        builderservice.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                    }
                    else
                    {
                        //this is for heading placed at the estimated efforts
                        if (r != Convert.ToInt32(splitservice[0]) - 2)
                        {
                            builderservice.Font.Bold = false;
                            builderservice.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                        }
                    }

                    for (int c = 0; c <= 3; c++)
                    {
                        builderservice.InsertCell();
                        builderservice.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builderservice.RowFormat.Alignment = RowAlignment.Center;

                        if (r == Convert.ToInt32(splitservice[0]) - 2)
                        {
                            builderservice.Font.Bold = true;
                            builderservice.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                        }

                        if (c == 0)
                        {
                            builderservice.CellFormat.Width = 50.0;
                        }
                        else
                        {
                            builderservice.CellFormat.Width = 200.0;
                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builderservice.Write(Temp);
                        Temp = "";
                    }
                    builderservice.EndRow();
                }
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }
        }

        public void Insertpayment()
        {
            try
            {
                int RowCount = myTable.Rows.Count;
                int ColumnCount = myTable.Columns.Count;
                builderpayment = new DocumentBuilder(doc);
                builderpayment.MoveToMergeField("Payments");
                //format of table of efforts///
                formattable(builderpayment);
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                int r = Convert.ToInt32(splitpayment[0]) - 2;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    DataArray[r, c] = myTable.Columns[c].ColumnName;
                    for (r = Convert.ToInt32(splitpayment[0]) - 2; r <= Convert.ToInt32(splitpayment[1]) - 2; r++)
                    {
                        DataArray[r, c] = myTable.Rows[r][c];

                    }
                }
                String Temp = "";
                for (r = Convert.ToInt32(splitpayment[0]) - 2; r <= Convert.ToInt32(splitpayment[1]) - 2; r++)
                {
                    if (r % 2 == 0)
                    {
                        builderpayment.Font.Bold = false;
                        builderpayment.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                    }
                    else
                    {
                        //this is for heading placed at the estimated efforts
                        if (r != Convert.ToInt32(splitservice[0]) - 2)
                        {
                            builderpayment.Font.Bold = false;
                            builderpayment.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
                        }
                    }

                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        builderpayment.InsertCell();
                        builderpayment.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builderpayment.RowFormat.Alignment = RowAlignment.Center;

                        if (r == Convert.ToInt32(splitservice[0]) - 2)
                        {
                            builderpayment.Font.Bold = true;
                            builderpayment.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Silver;
                        }

                        //formatting cell of payment rate//
                        if (c == 0 || c == 2)
                        {
                            builderpayment.CellFormat.Width = 50.0;
                        }
                        else
                        {
                            builderpayment.CellFormat.Width = 200.0;
                        }
                        Temp = Temp + DataArray[r, c] + "\t";
                        builderpayment.Write(Temp);
                        Temp = "";
                    }
                    builderpayment.EndRow();
                }
            }

            catch
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the  Spreadsheet and Template data worksheet');", true);
            }
        }

        public void formattable(DocumentBuilder buildersample)
        {

            buildersample.StartTable();

            ParagraphFormat paragraphFormat = buildersample.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 2;
            paragraphFormat.Alignment = ParagraphAlignment.Left;
            paragraphFormat.TabStops.Add(new TabStop(0));
            paragraphFormat.LeftIndent = 0;
            paragraphFormat.RightIndent = 0;
            paragraphFormat.SpaceAfter = 0;

            buildersample.RowFormat.Height = 15;
            //to get the rows and column in tabular form//
            buildersample.CellFormat.Borders.LineStyle = LineStyle.Single;
            buildersample.RowFormat.HeightRule = HeightRule.AtLeast;
            // Some special features for the header row.
            buildersample.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.FromArgb(198, 217, 241);
            buildersample.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            buildersample.Font.Size = 10;
            buildersample.Font.Name = "Arial";
            buildersample.Font.Bold = false;
            buildersample.CellFormat.Width = 100.0;
            buildersample.RowFormat.Height = 30.0;
            buildersample.RowFormat.HeightRule = HeightRule.Auto;
        }

        public void CreateHeaderAndFooter()
        {
            string primaryHeaderText = string.Empty;
            string secondaryHeaderText = string.Empty;
            string pageFooterText = string.Empty;

            primaryHeaderText = myTable.Rows.Cast<DataRow>().Where(r => r.Field<string>("field").ToLower().Trim() == "primaryheadertext").Select(r => r.Field<string>("value")).SingleOrDefault();
            secondaryHeaderText = myTable.Rows.Cast<DataRow>().Where(r => r.Field<string>("field").ToLower().Trim() == "secondaryheadertext").Select(r => r.Field<string>("value")).SingleOrDefault();
            pageFooterText = myTable.Rows.Cast<DataRow>().Where(r => r.Field<string>("field").ToLower().Trim() == "pagefootertext").Select(r => r.Field<string>("value")).SingleOrDefault();

            DocumentBuilder builder = new DocumentBuilder(doc);
            //Aspose.Words.Section
            Aspose.Words.Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;

            // Specify if we want headers/footers of the first page to be different from other pages.
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            // different headers/footers for odd and even pages.
            pageSetup.DifferentFirstPageHeaderFooter = true;

            pageSetup.BorderSurroundsHeader = true;
            pageSetup.HeaderDistance = 10;
            pageSetup.FooterDistance = 10;

            // Set font properties for header text.
            builder.Font.Color = System.Drawing.Color.Black;
            builder.Font.Name = "Calibri";
            builder.Font.Bold = true;
            builder.Font.Size = 12;

            //if (!string.IsNullOrEmpty(primaryHeaderText))
            //{
            //    // --- Create header for the first page. ---

            //    builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            //    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            //    // Set font properties for header text.
            //    builder.Font.Name = "Arial";
            //    builder.Font.Bold = true;
            //    builder.Font.Size = 12;
            //    // Specify header title for the first page.
            //    //builder.Write("Primary Header - TITLE PAGE");
            //    builder.Write(primaryHeaderText);
            //}

            if (!string.IsNullOrEmpty(secondaryHeaderText))
            {
                // --- Create header for pages other than first. ---
                pageSetup.HeaderDistance = 10;
                builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                builder.Write(secondaryHeaderText);
            }

            // --- Create footer for pages other than first. ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use table with two cells to make one part of the text on the line (with page numbering)
            // to be aligned left, and the other part of the text (with copyright) to be aligned right.
            Aspose.Words.Tables.Table table = builder.StartTable();

            // Clear table borders.
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            // Set first cell to 1/3 of the page width.
            //  builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            builder.CellFormat.Width = builder.PageSetup.PageWidth / 3;

            // Insert page numbering text here.
            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.

            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            builder.CellFormat.Borders.Top.ClearFormatting();
            builder.CellFormat.Borders.Top.Color = System.Drawing.Color.Black;
            builder.CellFormat.Borders.Top.DistanceFromText = 1;
            builder.CellFormat.Borders.Top.LineStyle = LineStyle.Hairline;

            //builder.CellFormat.Borders.Top.Shadow = true;

            builder.CellFormat.TopPadding = 10;
            // Align this text to the left.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (!string.IsNullOrEmpty(pageFooterText))
            {
                builder.InsertCell();
                builder.CellFormat.Width = builder.PageSetup.PageWidth * 2 / 3;

                builder.Write(pageFooterText);
            }

            // Align this text to the right.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();
            // Make page break to create a second page on which the primary headers/footers will be seen.

            builder.InsertBreak(BreakType.PageBreak);

            // Make section break to create a third page with different page orientation.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Get the new section and its page setup.
            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;

            // Set page orientation of the new section to landscape.
            // pageSetup.Orientation = Orientation.Landscape;

            // This section does not need different first page header/footer.
            // We need only one title page in the document and the header/footer for this page
            // has already been defined in the previous section
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section by default.
            // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
            // Page width is different for the new section and therefore we need to set 
            // a different cell widths for a footer table.
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section 
            // but with some minor modifications then it may be expedient to copy headers/footers
            // from the previous section and apply the necessary modifications where we want them.
            CopyHeadersFootersFromPreviousSection(currentSection);

            // Find the footer that we want to change.
            Aspose.Words.HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Aspose.Words.Tables.Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.Width = builder.PageSetup.PageWidth / 3;
            row.LastCell.CellFormat.Width = builder.PageSetup.PageWidth * 2 / 3;

        }

        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private static void CopyHeadersFootersFromPreviousSection(Aspose.Words.Section section)
        {
            Aspose.Words.Section previousSection = (Aspose.Words.Section)section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (Aspose.Words.HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }

        protected void Button5_Click(object sender, EventArgs e)
        {
            //if (listView1.SelectedIndex > -1 && listView2.SelectedIndex > -1 && listview3.SelectedIndex > -1)
            //{
            //    synchtemplate(myTable);
            //}
            //else
            //{
            //    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Select the Worksheet');", true);
            //}
        }

        private bool ValidateSpreadSheet()
        {
            bool response = false;
            if (myTable.Columns.Contains("field") && myTable.Columns.Contains("value"))
                response = true;
            return response;
        }
    }
}






