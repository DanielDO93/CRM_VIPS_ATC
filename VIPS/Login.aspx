<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Masterpages/Login.Master" CodeBehind="Login.aspx.vb" Inherits="VIPS.Login1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">

    <div id="site_content">

      <div class="content">
        <h1>Iniciar Sesión</h1>
        <div class="content_item">

  <asp:Login  ID="Login2" runat="server" BorderPadding="0" FailureText="El usuario o la contraseña son incorrectos." Height="100px" LoginButtonText="Login" RememberMeText="Recuerdame la proxima vez." TitleText="" UserNameLabelText="User:" Width="245px" style="margin-left:39%" ForeColor="Black" PasswordLabelText="Pass:" DisplayRememberMe="false">
             <CheckBoxStyle Font-Bold="True" Font-Names="Arial" Font-Size="XX-Small" />
             <InstructionTextStyle Font-Bold="True" Font-Names="Arial" />
             <LabelStyle  CssClass="LoginBox"  BorderStyle="None"/>
             <FailureTextStyle Font-Bold="True" Font-Size="XX-Small" />
             <LoginButtonStyle CssClass="Button" Height="20px" Width="50px" />
             <TextBoxStyle Width="120px" CssClass="LoginBox"/>
            </asp:Login>
      </div>
      </div>
    </div>
</asp:Content>
