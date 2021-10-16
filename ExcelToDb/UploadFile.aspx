<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadFile.aspx.cs" Inherits="ExcelToDb.UploadFile" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 61%;
            height: 202px;
        }
        .auto-style2 {
            width: 159px;
        }
        .auto-style3 {
            margin-left: 280px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
        </div>
        <table border="1" class="auto-style1" style="background-color: #FFFFCC">
            <tr>
                <td class="auto-style2">Upload Excel File</td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="auto-style2">&nbsp;</td>
                <td>
                    <asp:Button ID="btnSave" runat="server" BackColor="#FFCC00" OnClick="btnSave_Click" Text="Upload and Save To SqlServer" />
                </td>
            </tr>
        </table>
        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
        <p class="auto-style3">
            <asp:Button ID="btnView" runat="server" OnClick="btnView_Click" Text="View" />
        </p>
    </form>
</body>
</html>
