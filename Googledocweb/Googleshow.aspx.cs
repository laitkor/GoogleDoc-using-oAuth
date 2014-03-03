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


namespace Googledocweb
{
    public partial class Googleshow : System.Web.UI.Page
    {
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
        private bool Resizing = false;
        DataTable myTable;      
        static Assembly g_assembly;
        static DocX g_document;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void button1_Click(object sender, EventArgs e)
        {


            SpreadsheetsService GoogleExcelService;
            GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
            GoogleExcelService.setUserCredentials("abhishek.dubey@laitkor.com", "Lucknow@");
            SpreadsheetQuery query = new SpreadsheetQuery();
            SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
            //if (myFeed == null)
            //{
            //    label4.Text = "Invalid Credentials";
            //}
            foreach (SpreadsheetEntry mySpread in myFeed.Entries)
            {
                //string[] row = { mySpread.Title.Text, mySpread.Summary.Text, mySpread.Updated.ToShortDateString() };
                System.Web.UI.WebControls.ListItem listItem = new System.Web.UI.WebControls.ListItem(mySpread.Title.Text);
                listView1.Items.Add(listItem);



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
                    return;
                }

                SpreadsheetsService GoogleExcelService;
                GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-App");
                GoogleExcelService.setUserCredentials("abhishek.dubey@laitkor.com", "Lucknow@");
                // ListQuery query = new ListQuery("0AmYgMIof-5mgdGM2OGxoTmUyc3JRTFlMZ1BTUG5SOVE", "1", "public", "values");            
                // ListQuery query = new ListQuery("https://docs.google.com/a/laitkor.com/spreadsheet/ccc?key=0AttN4WWVg0qodF9RSG8tOXptV0RwZm1LOWFJQ3g0Mnc#gid=0", "1", "public", "values");
                //ListFeed myFeed = GoogleExcelService.Query(query);
                SpreadsheetQuery query = new SpreadsheetQuery();
                SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
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
                            }
                        }
                    }
                }
            }
            //System.Windows.Forms.MessageBox.Show("Data Reading is Completed");
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Data Reading is Completed');", true);
        }

        protected void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedIndex > -1)
                spreadsheetName = listView1.SelectedItem.Text;
            else
                return;
            if (spreadsheetName.Length > 0)
            {
                SpreadsheetsService GoogleExcelService;
                GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-Test-App");
                GoogleExcelService.setUserCredentials("abhishek.dubey@laitkor.com", "Lucknow@");
                SpreadsheetQuery query = new SpreadsheetQuery();
                SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
                foreach (SpreadsheetEntry mySpread in myFeed.Entries)
                {
                    if (mySpread.Title.Text == spreadsheetName)
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

        

        protected void Button3_Click(object sender, EventArgs e)
        {

            //Response.ContentType = "application/pdf";
            //Response.AddHeader("content0disposition",
            //"attachment;filename=UserDetails.pdf");
            //Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //StringWriter sw = new StringWriter();
            //HtmlTextWriter hw = new HtmlTextWriter(sw);
            //dataGridView1.AllowPaging = false;
            //dataGridView1.DataBind();
            //dataGridView1.RenderControl(hw);
            //dataGridView1.HeaderRow.Style.Add("width", "15%");
            //dataGridView1.HeaderRow.Style.Add("font-size", "10px");
            //dataGridView1.Style.Add("text-decoration", "none");
            //dataGridView1.Style.Add("font-family", "Arial, Helvetica, sans-serif");
            //dataGridView1.Style.Add("font-size", "8px");
            //StringReader sr = new StringReader(sw.ToString());
            //StreamReader reader = new StreamReader(new MemoryStream(Encoding.ASCII.GetBytes(sw.ToString())));
            //iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A2, 7f, 7f, 7f, 0f);
            //HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            //PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            //pdfDoc.Open();
            //htmlparser.Parse(reader);
            //pdfDoc.Close();
            //Response.Write(pdfDoc);
            //Response.End(); 

            //Here set page size as A4


            //Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=Export.pdf");
            //Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //StringWriter sw = new StringWriter();
            //HtmlTextWriter hw = new HtmlTextWriter(sw);
            //HtmlForm frm = new HtmlForm();
            //dataGridView1.Parent.Controls.Add(frm);
            //frm.Attributes["runat"] = "server";
            //frm.Controls.Add(dataGridView1);
            //frm.RenderControl(hw);
            //StringReader sr = new StringReader(sw.ToString());
            //iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);
            //HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            //PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            //pdfDoc.Open();
            //htmlparser.Parse(sr);
            //pdfDoc.Close();
            //Response.Write(pdfDoc);
            //Response.End();
            createpdf();
        }


        public void bindgrid()
        {
            SpreadsheetsService GoogleExcelService;
            GoogleExcelService = new SpreadsheetsService("Spreadsheet-Abhishek-App");
            GoogleExcelService.setUserCredentials("abhishek.dubey@laitkor.com", "Lucknow@");
            SpreadsheetQuery query = new SpreadsheetQuery();
            SpreadsheetFeed myFeed = GoogleExcelService.Query(query);
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
                        }
                    }
                }
            }
        }
        

        public void createpdf()
        {
            bindgrid();
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

            //PdfPCell PdfPCellhrs = new PdfPCell(new Phrase(new Chunk("Payment Information", font9)));
            //PdfPCellhrs.Colspan = 6;
            //PdfPCellhrs.Border = 0;
            //PdfPCellhrs.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            //PdfTable.AddCell(PdfPCellhrs);
           // DataSet dsttlhrs = gethrs();
            //str = "Total Working Hours:" + " " + dsttlhrs.Tables[0].Rows[0]["WorkingHours"].ToString() + "Hours";
            //PdfPCell PdfPCelltotal = new PdfPCell(new Phrase(new Chunk(str, font8)));
            //PdfPCelltotal.Colspan = 6;
            //PdfPCelltotal.Border = 0;
            //PdfTable.AddCell(PdfPCelltotal);

            //str = "Total Cost:" + " " + "$" + dsttlhrs.Tables[0].Rows[0]["TotalCost"].ToString();
            //PdfPCell PdfPCellProcost = new PdfPCell(new Phrase(new Chunk(str, font8)));
            //PdfPCellProcost.Colspan = 6;
            //PdfPCellProcost.Border = 0;
            //PdfTable.AddCell(PdfPCellProcost);
            //PdfTable.AddCell(PdfPCellBlank);


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
            //PdfPCell PdfPCellSign = new PdfPCell(new Phrase(new Chunk("", font9)));
            //PdfPCellSign.Colspan = 6;
            //PdfPCellSign.Border = 0;
            //PdfPCellSign.BackgroundColor = new iTextSharp.text.Color(202, 197, 180);
            //PdfTable.AddCell(PdfPCellSign);

            //string imagepath = Server.MapPath("Image");
            //string URL = System.Web.HttpContext.Current.Request.Url.AbsoluteUri.ToString();
            //var fileNamesign = "sign.jpg";
            //var fullPathsign = imagepath + "\\" + fileNamesign;
            //iTextSharp.text.Image jpgsign = iTextSharp.text.Image.GetInstance(fullPathsign);
            //PdfPCell pdfpcellimagesign = new PdfPCell(jpgsign);
            //pdfpcellimagesign.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
            //pdfpcellimagesign.Colspan = 6;
            //pdfpcellimagesign.Border = 0;
            //PdfTable.AddCell(pdfpcellimagesign);

            #endregion sign

            doc.Add(PdfTable);
            //Chunk myFooter = new Chunk("Page " + (doc.PageNumber), FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8));
            //PdfPCell footer = new PdfPCell(new Phrase(myFooter));
            //footer.Border = Rectangle.NO_BORDER;
            //footer.HorizontalAlignment = Element.ALIGN_CENTER;
            //PdfTable.AddCell(footer); 

            #region footer
            Chunk myFooter = new Chunk(DateTime.Now.ToShortDateString() + "      " + "Page " + (doc.PageNumber + 1),
               FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8, new iTextSharp.text.Color(46, 84, 141)));
            HeaderFooter footer = new HeaderFooter(new Phrase(myFooter), false);    //Create a footer object with the chunk data
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
            // proc.Start();
            //   Response.TransmitFile(fullPath);
           // Button2.Enabled = false;
            Response.Flush();
            Response.End();

        }

       



        // Create an invoice for a factitious company called "The Happy Builder".
      


        // Create an invoice template.
       

     

        protected void Button5_Click(object sender, EventArgs e)
        {    

            //code to export gridview to word file/////

            //Response.AddHeader("content-disposition", "attachment;filename=myreport.doc");
            //Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.ContentType = "application/vnd.word";
            //System.IO.StringWriter swriter = new System.IO.StringWriter();
            //System.Web.UI.HtmlTextWriter htmlwriter = new HtmlTextWriter(swriter);
            //// Create a form to contain the gridview(MyGridView)
            //HtmlForm mynewform = new HtmlForm();
            //dataGridView1.Parent.Controls.Add(mynewform);
            //mynewform.Attributes["runat"] = "server";
            //mynewform.Controls.Add(dataGridView1);
            //mynewform.RenderControl(htmlwriter);
            //Response.Write(swriter.ToString());
            //Response.End();

            bindgrid();
            //method to convert datatable to word file////
            CreateWordTableWithDataTable(myTable);

        }



        public void CreateWordTableWithDataTable(DataTable myTable )
        {
            int RowCount = myTable.Rows.Count; 
            int ColumnCount = myTable.Columns.Count;
            Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
            int r = 0;
            for (int c = 0; c <= ColumnCount-1; c++)
            {
                DataArray[r, c] = myTable.Columns[c].ColumnName;
                for (r = 0; r <=RowCount-1; r++)
                {
                    DataArray[r, c] = myTable.Rows[r][c];
                } //end row loop
            } //end column loop

            Microsoft.Office.Interop.Word.Document Doc = new Microsoft.Office.Interop.Word.Document();
            Doc.Application.Visible = true;
            Doc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;         
            dynamic Range = Doc.Content.Application.Selection.Range;
            String Temp ="";
            for (r = 0; r <= RowCount-1; r++)
            {
                for (int c = 0; c <=ColumnCount-1; c++)
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
            Range.ConvertToTable(ref Separator,ref RowCount, ref ColumnCount, Type.Missing, ref Format,ref ApplyBorders, Type.Missing, Type.Missing, Type.Missing,
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
            for (int c = 0; c <= ColumnCount-1; c++)
            {
                Doc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = myTable.Columns[c].ColumnName;
            }

            Doc.Application.Selection.Tables[1].Rows[1].Select();
            Doc.Application.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                      
                    }

        protected void listView2_SelectedIndexChanged1(object sender, EventArgs e)
        {
            wrkSheetName = "";
            if (listView2.SelectedIndex > -1)
            {
                wrkSheetName = listView2.SelectedItem.Text;

            }
            else
            {
                return;
            }
        }

        
    }
}