<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Masterpages/Login.Master" CodeBehind="Recuperar_Password.aspx.vb" Inherits="VIPS.Recuperar_Password" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .auto-style3 {
            text-align: left;
        }

        .auto-style6 {
            width: 61px;
        }

        .auto-style10 {
            text-align: right;
            width: 348px;
        }

        .auto-style11 {
            text-align: right;
            width: 323px;
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




                    <asp:Panel ID="Panel1" runat="server"  CssClass="Panel">

                        <h2>
                            <asp:Label ID="Label1" runat="server" Text="Cambio de Contraseña"></asp:Label></h2>



                        <table style="width: 100%;" class="Buscador">

                            <tr>
                                <td class="auto-style6"></td>
                                <td class="auto-style11">Usuario:</td>
                                <td class="auto-style3">
                                    <asp:TextBox ID="TextBox1" runat="server" Width="148px" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                </td>
                                <td></td>
                            </tr>
                            <tr>
                                <td class="auto-style6"></td>
                                <td class="auto-style11">Fecha de Nacimiento (dd-mm-aaaa):</td>
                                <td class="auto-style3">
                                    <asp:TextBox ID="TextBox2" runat="server" Width="40px" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="TextBox2" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                    <asp:TextBox ID="TextBox3" runat="server" Width="40px" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="TextBox3" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                    <asp:TextBox ID="TextBox4" runat="server" Width="40px" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox4" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                </td>
                                <td></td>
                            </tr>
                            <tr>
                                <td class="auto-style6"></td>
                                <td class="auto-style11">Nuevo Password:</td>
                                <td class="auto-style3">
                                    <asp:TextBox ID="TextBox5" runat="server" Width="148px" TextMode="Password" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="TextBox5" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                </td>
                                <td></td>
                            </tr>
                            <tr>
                                <td class="auto-style6"></td>
                                <td class="auto-style11">Confirma Password:</td>
                                <td class="auto-style3">
                                    <asp:TextBox ID="TextBox6" runat="server" Width="148px" TextMode="Password" CssClass="textos"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="TextBox6" ErrorMessage="*" ForeColor="Red" ValidationGroup="Pass"></asp:RequiredFieldValidator>
                                </td>
                                <td></td>
                            </tr>
                            <tr>
                                <td class="auto-style6"></td>
                                <td class="textos" colspan="2">
                                    <asp:Button ID="Button2" runat="server" Font-Bold="True" Font-Size="X-Small" Text="Guardar" ValidationGroup="Pass" ccsclass="Button" CssClass="Button" Height="20" Width="80" />
                                </td>
                            </tr>

                        </table>


                    </asp:Panel>





                </ContentTemplate>
            </asp:UpdatePanel>


        </div>
    </div>


</asp:Content>
