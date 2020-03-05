<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Masterpages/Content.Master" CodeBehind="Password.aspx.vb" Inherits="VIPS.Password" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
    .auto-style3 {
        text-align: left;
    }
        .auto-style11 {
            text-align: right;
            width: 368px;
        }
        .auto-style12 {
            text-align: right;
            width: 365px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div id="site_content">
    <div class="content">
        <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
            </asp:ToolkitScriptManager>
        <h1>Cambio de Contraseña</h1>

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
          <ContentTemplate>

 <div style="margin-left: 2%;">

        <asp:Panel ID="Panel1" runat="server" Width="896px" height="105px" CssClass="Panel">
            <h2>Cambio de Contraseña</h2>
            <div style="width:888px; margin-left: 4px; margin-right:4px;">

                <table style="width:100%;">

                    <tr>
                        <td></td>
                        <td class="auto-style12">Nuevo Password:</td>
                        <td class="auto-style3">
                            <asp:TextBox ID="TextBox1" runat="server" TextMode="Password"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                        </td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td class="auto-style12">Confirma Password:</td>
                        <td class="auto-style3">
                            <asp:TextBox ID="TextBox2" runat="server" TextMode="Password"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                        </td>
                        <td></td>
                    </tr>
                   <tr>
                        <td></td>
                        <td class="textos" colspan="2">
                            <asp:Button ID="Button2" runat="server" Font-Bold="True" Font-Size="X-Small" Text="Guardar" ValidationGroup="Pass" ccsclass="Button" CssClass="Button" Height="20" Width="80" />
                        </td>
                        <td></td>
                    </tr>

                </table>

              </div>
        </asp:Panel>

     </div>


              <br />
              <br />

          </ContentTemplate>
      </asp:UpdatePanel>


        </div>
        </div>


</asp:Content>
