<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MainForm.aspx.cs" Inherits="Googledocweb.MainForm" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Cloud Merge :: Generate Proposal</title>
    <script src="js/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.1/jquery.min.js" type="text/javascript"></script>
    <link href="Styles/Site.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style1
        {
            width: 184px;
        }
        .style2
        {
            width: 219px;
        }
        .style3
        {
            width: 184px;
            height: 263px;
        }
    </style>
    <script type="text/javascript">
        $(document).ready(function () {

            OpenAuthURL();

        });

        function OpenAuthURL() {
            var URL = $("#<%= hurl.ClientID%>").val();
            if (URL != "") {
                //window.open(URL, '_blank');
            }
        }

       
    </script>
    <script type="text/javascript">
        function MyEndRequest() {
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <table cellpadding="0" cellspacing="0" border="0" align="center" style="margin-top: 10px"
        width="100%">
        <tr>
            <td align="center">
                <div style="background-image: url('../Image/logo2.png'); height: 81px; margin: 0 auto;
                    background-repeat: no-repeat; background-position: center">
                    <asp:Label ID="Label5" runat="server" Style="color: Teal; font-family: Arial; font-weight: 900"
                        class="lblheading" Text=""></asp:Label>
                </div>
                <asp:Panel ID="Panel_Main" runat="server" BorderStyle="None">
                    <div class="selectorlogin" style="background-color: White; width: 850px; height: 395px;
                        -moz-border-radius: 10px; border-radius: 10px; -webkit-border-radius: 10px; background-position: center;
                        box-shadow: 1px 1px 2px rgba(0,0,0,0.05); padding-left: 0px; padding-top: 0px;
                        padding: 10px; margin-top: 20px; margin-bottom: 20px; background-color: white;
                        border: 1px solid #babbbb">
                        <asp:UpdatePanel ID="Update_Main" runat="server">
                            <ContentTemplate>
                                <table align="center" style="margin-top: 10px; padding-top: 30px; height: 148%; width: 89%;">
                                    <tbody valign="middle">
                                        <tr>
                                            <td colspan="4">
                                                <asp:Label ID="Lbl_Msg" runat="server" Text=""></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="style2">
                                                <%--<asp:TextBox ID="TextBox2" runat="server" Style="margin-left: 0px" Width="200px"></asp:TextBox>--%>
                                            </td>
                                            <td>
                                                &nbsp;
                                            </td>
                                            <td colspan="2">
                                                <asp:HiddenField ID="hurl" runat="server" Value="0" />
                                            </td>
                                        </tr>
                                        <tr class="height40">
                                            <td class="style1">
                                                <asp:Label CssClass="fntclrlogin" ID="Label3" runat="server" Text="Spreadsheet" Style="margin-top: 14px;
                                                    padding-left: 90px" ForeColor="#4c4c4c" Font-Bold="True" Font-Names="Open Sans"
                                                    Font-Size="12pt"></asp:Label>
                                            </td>
                                            <td class="style2">
                                                <%--<asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                                <ContentTemplate>--%>
                                                <asp:ListBox ID="listView1" runat="server" AutoPostBack="True" Style="margin-top: 14px"
                                                    CssClass="roundList" OnSelectedIndexChanged="listView1_SelectedIndexChanged"
                                                    Width="230px"></asp:ListBox>
                                                <%-- </ContentTemplate>
                                            </asp:UpdatePanel>--%>
                                            </td>
                                            <td class="field required">
                                                <asp:Label CssClass="fntclrlogin" ID="Label1" runat="server" Text="Worksheet" ForeColor="#4c4c4c"
                                                    Font-Bold="True" Font-Names="Open Sans" Style="padding-left: 10px" Font-Size="12pt"></asp:Label>
                                            </td>
                                            <td>
                                                <%-- <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                                <ContentTemplate>--%>
                                                <asp:ListBox ID="listView2" runat="server" Width="230px" Style="margin-top: 14px"
                                                    CssClass="roundList" OnSelectedIndexChanged="listView2_SelectedIndexChanged">
                                                </asp:ListBox>
                                                <%-- </ContentTemplate>
                                            </asp:UpdatePanel>--%>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" style="padding-left: 224px">
                                                <asp:Button ID="button2" runat="server" Text="Export To PDF" Style="margin-right: 20px;
                                                    margin-top: 20px" OnClick="button2_Click" BackColor="#999966" Visible="False" />
                                                <asp:Button ID="Button4" runat="server" CssClass="btn_Export" OnClick="Button4_Click1"
                                                    Height="37px" Width="149px" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4">
                                                <asp:UpdateProgress ID="Progressbar1" runat="server">
                                                    <ProgressTemplate>
                                                        <div>
                                                            <table class="progressBarTable" cellspacing='2' cellpadding='2'>
                                                                <tr>
                                                                    <td>
                                                                        <img src="Image/updateprogress.gif" />
                                                                    </td>
                                                                    <td>
                                                                        <span>Please wait while your request is being processed.</span>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </div>
                                                    </ProgressTemplate>
                                                </asp:UpdateProgress>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" class="">
                                                <asp:GridView ID="dataGridView1" runat="server" Width="167px" Visible="False" Height="1px">
                                                </asp:GridView>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="Button4" />
                            </Triggers>
                        </asp:UpdatePanel>
                    </div>
                    </div>
                </asp:Panel>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
