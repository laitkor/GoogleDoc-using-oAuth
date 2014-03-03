<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Googleshow.aspx.cs" Inherits="Googledocweb.Googleshow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table align="center">
    <tr>
    <td>
    <%--<asp:ListView ID="listView1" runat="server" 
            onselectedindexchanged="listView1_SelectedIndexChanged">
        </asp:ListView>--%>
    <asp:ListBox ID="listView1" runat="server" AutoPostBack="True" 
            onselectedindexchanged="listView1_SelectedIndexChanged" Width="180px"></asp:ListBox>
    </td>
    </tr>
    <tr>
    <td>
    <%--<asp:ListView ID="listView2" runat="server" 
            onselectedindexchanged="listView2_SelectedIndexChanged">
        </asp:ListView>--%>
        <asp:ListBox ID="listView2" runat="server" Width="180px" 
            onselectedindexchanged="listView2_SelectedIndexChanged1"></asp:ListBox>
    </td>
    </tr>
    <tr>
    <td>
     <asp:Button ID="button1" runat="server" Text="Pull Spread sheet" 
            onclick="button1_Click" />
        <asp:Button ID="button2" runat="server" Text="Read Data" 
            onclick="button2_Click" />
    </td>
    <td>
        <asp:Button ID="Button3" runat="server" onclick="Button3_Click" 
            Text="Export Pdf" />
        <asp:Button ID="Button4" runat="server" onclick="Button4_Click" 
            Text="Export To Word" />
        <asp:Button ID="Button5" runat="server" onclick="Button5_Click" Text="Button" />
    </td>
    </tr>
    <tr>
    <td>
        <asp:GridView ID="dataGridView1" runat="server" Width="243px" Visible="False">
        </asp:GridView>
    </td>
    </tr>
    </table>
        

       
    
    </div>
    
    </form>
</body>
</html>
