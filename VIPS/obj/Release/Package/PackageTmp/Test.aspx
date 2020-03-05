<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Test.aspx.vb" Inherits="VIPS.Test" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>

<body>
    <form id="form1" runat="server">

        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

        <div>

            <asp:TextBox ID="txtUrl" runat="server" Text=""></asp:TextBox>
            <asp:Button ID="Button1" Text="Capture" runat="server" OnClick="Capture" />
            <br />
            <asp:Image ID="imgScreenShot" runat="server" Height="500" Width="700" Visible="false" />
        </div>



        <div style="border-style: dotted; float: left; width: 500px; height: 500px; background-color: #FFFFFF;">
            <iframe id="mainiFrame" name="mainiFrame" scrolling="auto" frameborder="2"
                height="400px" width="500px"
                src="https://callcentervips.teamolo.info"></iframe>
        </div>

    </form>
</body>
</html>
