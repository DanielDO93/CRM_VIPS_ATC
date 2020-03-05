<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Masterpages/Inbound.Master" CodeBehind="Tipificacion.aspx.vb" Inherits="VIPS.Inbound" EnableEventValidation="false" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>


<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">



    <link href="../CSS/ValidationEngine.css" rel="stylesheet" type="text/css" />
    <script src="../JS/jquery.validationEngine-en.js" charset="utf-8"></script>
    <script src="../JS/jquery.validationEngine.js" charset="utf-8"></script>
    <script id="grid" type="text/javascript">


        $(function () {

            if (document.getElementById('ContentPlaceHolder1_HiddenField2').value == 1) {
                $('#miModal').attr('id', 'the_new_id');
            };


        })




        function pageLoad() {
            jQuery("#form1").validationEngine();
        }




        function DateFormat(field, rules, i, options) {
            var regex = /^(0?[1-9]|[12][0-9]|3[01])[\/\-](0?[1-9]|1[012])[\/\-]\d{4}$/;
            if (!regex.test(field.val())) {
                return "La fecha debe estar en formato DD/MM/AAAA"
            }
        }



        function On(GridView) {
            if (GridView != null) {
                GridView.originalBgColor = GridView.style.backgroundColor;
                GridView.style.backgroundColor = "#EEEEEE";
                GridView.style.Color = "#FFFFFF";
            }
        }

        function Off(GridView) {
            if (GridView != null) {
                GridView.style.backgroundColor = GridView.originalBgColor;
            }
        }



    </script>


</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">

    <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" AsyncPostBackTimeout="0" EnablePageMethods="True"></asp:ToolkitScriptManager>
    <asp:HiddenField ID="HiddenField2" runat="server" Value="1" />
    <asp:HiddenField ID="HiddenField3" runat="server" Value="1" />
    <asp:HiddenField ID="Cliente" runat="server" Value="x" />

    <div id="site_content">

        <div class="content">

            <h1>
                <asp:Label ID="Label75" runat="server" Text="Nueva Interacción" Font-Size="100%" ForeColor="#063E4C"></asp:Label></h1>

            <%-- **************************************** FORMULARIOS MODALES **************************************** --%>


            <div id="modalBuscador" class="modal">
                <asp:UpdatePanel ID="UpdatePanel4" runat="server">

                    <ContentTemplate>


                        <div class="modal-contenido2">
                            <h1>Busqueda de Clientes</h1>

                            <div class="form_default">
                                <ul>
                                    <li>
                                        <asp:Label ID="Label25" runat="server" Text="Nombre:"></asp:Label></li>
                                    <li>
                                        <asp:TextBox ID="TextBox17" runat="server" CssClass="textbox" AutoPostBack="True"></asp:TextBox>
                                    </li>
                                    <li>
                                        <asp:Label ID="Label64" runat="server" Text="Teléfono:"></asp:Label></li>
                                    <li>
                                        <asp:TextBox ID="TextBox51" runat="server" CssClass="textbox" AutoPostBack="True"></asp:TextBox>
                                    </li>
                                </ul>
                            </div>

                            <br />
                            <div class="search-table">
                                <asp:GridView ID="GridView2" runat="server"></asp:GridView>
                            </div>


                            <div class="form_default">
                                <div class="Button" style="width: 100px; margin-left: 250px;">
                                    <a href="#">Cerrar</a>
                                </div>
                            </div>
                            <br />
                        </div>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </div>


            <div id="modalError" class="modal">
                <div class="modal-contenido">
                    <h1>¡No Guardado!</h1>


                    <asp:Label ID="Label76" runat="server" Text="!El pedido debe estar asociado a un cliente!"></asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="Label78" runat="server" Text="Selecciona un cliente o da click en Nuevo Cliente"></asp:Label>
                    <br />
                    <br />

                    <div class="Button" style="width: 100px; margin-left: 150px;">
                        <a href="#">Cerrar</a>
                    </div>
                    <br />
                </div>
            </div>


            <div id="modalQueja" class="modal">
                <div class="modal-contenido">
                    <h1>¡Nueva Queja!</h1>


                    <asp:Label ID="Label30" runat="server" Text="Se ha generado una nueva queja con el numero de folio:"></asp:Label>
                    <br />
                    <br />
                    <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                        <ContentTemplate>
                            <asp:Label ID="Label31" runat="server" Text=""></asp:Label>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                    <br />
                    <div class="Button" style="width: 100px; margin-left: 150px;">
                        <a href="#" onclick="alv();">Cerrar</a>
                    </div>
                    <br />
                </div>
            </div>

            <div id="VOLO" class="modal">
                <div class="modal-contenido" style="margin-left: 250px; margin-top: 10px; width: 810px;">

                    <iframe id="iFrameID2" name="iFrameID" src="https://callcentervips.teamolo.info" scrolling="auto" height="550px" width="800px"></iframe>

                    <br />
                    <br />
                    <div class="Button" style="width: 100px; margin-left: 350px;">
                        <a href="#">Cerrar</a>
                    </div>
                    <br />
                </div>
            </div>

            <div id="miModal" class="modal">
                <div class="modal-contenido">
                    <h1>¡Atención!</h1>


                    <asp:Label ID="LabelModal1" runat="server" Text="¡Este telefono está asociado a mas de un registro!"></asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="LabelModal2" runat="server" Text="Selecciona el cliente que está llamando."></asp:Label>

                    <div class="modal-tabla">
                        <asp:GridView ID="GridView1" runat="server"></asp:GridView>
                    </div>


                    <div class="Button" style="width: 100px; margin-left: 150px;">
                        <a href="#" onclick="mostrar()">Buscar</a>
                    </div>
                    <br />
                    <div class="Button" style="width: 100px; margin-left: 150px;">
                        <a href="#">Cerrar</a>
                    </div>
                    <br />
                </div>
            </div>

            <%-- **************************************************************************************************** --%>

            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>

           
                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label22" runat="server" Text="Ultima Interacción:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox14" runat="server" CssClass="textbox" Enabled="False"></asp:TextBox></li>
                            <li>
                                <asp:Label ID="Label23" runat="server" Text="Ultimo Pedido:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox15" runat="server" CssClass="textbox" Enabled="False"></asp:TextBox></li>
                            <li>
                                <asp:Label ID="Label24" runat="server" Text="Ultima Queja:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox16" runat="server" CssClass="textbox" Enabled="False"></asp:TextBox></li>
                        </ul>
                    </div>


                    <div style="text-align: center; width: 100%; margin-left: 220px; display: none;">
                        <asp:HiddenField ID="HiddenField1" runat="server" Value="3" />
                        <div id="CL1" runat="server" class="C1"></div>
                        <div id="CL2" runat="server" class="C2"></div>
                        <div id="CL3" runat="server" class="C3"></div>
                        <div id="CL4" runat="server" class="C4"></div>
                        <div id="CL5" runat="server" class="C5"></div>
                    </div>



                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label1" runat="server" Text="Nombres:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox1" runat="server" CssClass="textbox validate[required]"></asp:TextBox></a></li>
                            <li>
                                <asp:Label ID="Label2" runat="server" Text="Apellido Paterno:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox2" runat="server" CssClass="textbox validate[required]"></asp:TextBox></li>
                            <li>
                                <asp:Label ID="Label11" runat="server" Text="Apellido Materno:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox4" runat="server" CssClass="textbox validate[required]"></asp:TextBox></li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label10" runat="server" Text="Teléfono Móvil:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox5" runat="server" CssClass="textbox validate[required,custom[integer],maxSize[10],minSize[10]]"></asp:TextBox></li>
                            <li>
                                <asp:Label ID="Label12" runat="server" Text="Teléfono Casa:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox6" runat="server" CssClass="textbox validate[required,custom[integer],maxSize[10],minSize[10]]"></asp:TextBox></li>
                            <li>
                                <asp:Label ID="Label13" runat="server" Text="Teléfono Oficina:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox7" runat="server" CssClass="textbox validate[required,custom[integer],maxSize[10],minSize[10]]"></asp:TextBox></li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label14" runat="server" Text="Calle:"></asp:Label></li>
                            <li style="width: 280px;">
                                <asp:TextBox ID="TextBox8" runat="server" CssClass="textbox validate[required]" Width="250px"></asp:TextBox></li>
                            <li style="width: 50px;">
                                <asp:Label ID="Label15" runat="server" Text="No. Ext:"></asp:Label></li>
                            <li style="width: 50px;">
                                <asp:TextBox ID="TextBox9" runat="server" CssClass="textbox validate[required]" Width="25px" TextMode="Number"></asp:TextBox></li>
                            <li style="width: 50px;">
                                <asp:Label ID="Label16" runat="server" Text="No. Int:"></asp:Label></li>
                            <li style="width: 50px;">
                                <asp:TextBox ID="TextBox10" runat="server" CssClass="textbox" Width="25px"></asp:TextBox></li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label3" runat="server" Text="Colonia:"></asp:Label></li>
                            <li>
                                <asp:DropDownList ID="DropDownList1" runat="server" CssClass="textbox validate[required]" AutoPostBack="True">
                                    <asp:ListItem Value="">-</asp:ListItem>
                                </asp:DropDownList>
                            </li>
                            <li>
                                <asp:Label ID="Label4" runat="server" Text="Delegación/Municipio:"></asp:Label></li>
                            <li>
                                <asp:DropDownList ID="DropDownList2" runat="server" CssClass="textbox validate[required]" AutoPostBack="True">
                                    <asp:ListItem Value="">-</asp:ListItem>
                                </asp:DropDownList>
                            </li>
                        </ul>
                    </div>
                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label5" runat="server" Text="Estado:"></asp:Label></li>
                            <li>
                                <asp:DropDownList ID="DropDownList3" runat="server" CssClass="textbox validate[required]" AutoPostBack="True">
                                </asp:DropDownList></li>
                            <li>
                                <asp:Label ID="Label6" runat="server" Text="Codigo Postal:"></asp:Label></li>
                            <li>
                                <asp:TextBox ID="TextBox3" runat="server" CssClass="textbox validate[required]" AutoPostBack="True"></asp:TextBox></li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li style="width: 80px; margin-left: -270px;">
                                <asp:Label ID="Label17" runat="server" Text="Referencias:"></asp:Label></li>
                            <li style="margin-left: -170px;">
                                <asp:TextBox ID="TextBox11" runat="server" CssClass="textbox" Width="450px"></asp:TextBox>
                            </li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li style="width: 80px; margin-left: -270px;">
                                <asp:Label ID="Label21" runat="server" Text="e-mail:"></asp:Label></li>
                            <li style="margin-left: -170px;">
                                <asp:TextBox ID="TextBox12" runat="server" CssClass="textbox" Width="250px"></asp:TextBox>
                            </li>
                        </ul>
                    </div>

  

                    <%-- ***************************************** TIPIFICACION ***************************************** --%>


                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label7" runat="server" Text="Tipificación 1:"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList4" runat="server" AutoPostBack="True" CssClass="textbox validate[required]">
                                </asp:DropDownList>
                            </li>
                            <li>
                                <asp:Label ID="Label8" runat="server" Text="Tipificación 2:" Visible="False"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True" CssClass="textbox validate[required]" Visible="False">
                                </asp:DropDownList>
                            </li>
                            <ul />
                    </div>

                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label9" runat="server" Text="Tipificación 3:" Visible="False"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList6" runat="server" AutoPostBack="True" CssClass="textbox validate[required]" Visible="False">
                                </asp:DropDownList>
                            </li>
                            <li>
                                <asp:Label ID="Label37" runat="server" Text="Tipificación 4:" Visible="False"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList10" runat="server" AutoPostBack="True" CssClass="textbox validate[required]" Visible="False">
                                </asp:DropDownList>
                            </li>
                        </ul>
                    </div>

                    <div class="form_default">
                        <ul>
                            <li>
                                <asp:Label ID="Label41" runat="server" Text="Tipificación 5:" Visible="False"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList11" runat="server" AutoPostBack="True" CssClass="textbox validate[required]" Visible="False">
                                </asp:DropDownList>
                            </li>
                            <li>
                                <asp:Label ID="Label42" runat="server" Text="Tipificación 6:" Visible="False"></asp:Label>
                            </li>
                            <li>
                                <asp:DropDownList ID="DropDownList12" runat="server" AutoPostBack="True" CssClass="textbox validate[required]" Visible="False">
                                    <asp:ListItem Value="">-Selecciona-</asp:ListItem>
                                    <asp:ListItem Value="Restaurante">Restaurante</asp:ListItem>
                                    <asp:ListItem Value="Repartidor">Repartidor</asp:ListItem>
                                    <asp:ListItem Value="Call Center">Call Center</asp:ListItem>

                                </asp:DropDownList>
                            </li>
                        </ul>
                    </div>

             

                    <%-- **************************************** QUEJAS  NUEVAS **************************************** --%>

                    <div class="form_default" runat="server" id="divQuejas" visible="false">

                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label57" runat="server" Text="Pedido:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox46" runat="server" CssClass="textbox" AutoPostBack="true"></asp:TextBox></a></li>
                                <li>
                                    <asp:Label ID="Label58" runat="server" Text="Restaurante:"></asp:Label></li>
                                <li>
                                    <asp:DropDownList ID="DropDownList14" runat="server" CssClass="textbox validate[required]"></asp:DropDownList></li>
                                <li>
                                    <asp:Label ID="Label59" runat="server" Text="Repartidor:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox48" runat="server" CssClass="textbox"></asp:TextBox></li>
                            </ul>
                        </div>

                        <div class="form_default" runat="server" id="divQuejas1" style="margin-left: -180px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label26" runat="server" Text="Descripción:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox18" runat="server" CssClass="textbox validate[required]" Width="450px"></asp:TextBox>
                                </li>
                            </ul>
                        </div>
                        <div class="form_default" runat="server" id="divQuejas3" style="margin-left: -180px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label28" runat="server" Text="Solución:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox20" runat="server" CssClass="textbox validate[required]" Width="450px"></asp:TextBox>
                                </li>
                            </ul>
                        </div>


                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label29" runat="server" Text="Status Queja:"></asp:Label>
                                </li>
                                <li>
                                    <asp:DropDownList ID="DropDownList9" runat="server" AutoPostBack="True" CssClass="textbox validate[required]">
                                        <asp:ListItem Value="">-Selecciona-</asp:ListItem>
                                        <asp:ListItem Value="1">Abierta</asp:ListItem>
                                        <asp:ListItem Value="0">Cerrada</asp:ListItem>
                                    </asp:DropDownList>
                                </li>
                            </ul>
                        </div>
                    </div>

                    <%-- ************************************** SEGUIMIENTO QUEJAS ************************************** --%>

                    <div class="form_defult" runat="server" id="SeguimientoQuejas" visible="false">
                        <h2>Quejas Abiertas</h2>
                        <div class="quejas-abiertas">
                            <asp:GridView ID="GridView3" runat="server">
                            </asp:GridView>
                        </div>
                        <h2>Buscador de Quejas</h2>


                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label32" runat="server" Text="No. Queja:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox21" runat="server" CssClass="textbox" AutoPostBack="true"></asp:TextBox></a></li>
                                <li>
                                    <asp:Label ID="Label33" runat="server" Text="Nombre:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox22" runat="server" CssClass="textbox" AutoPostBack="true"></asp:TextBox></li>
                            </ul>
                        </div>
                        <div class="quejas-abiertas">
                            <asp:GridView ID="GridView4" runat="server">
                            </asp:GridView>
                        </div>
                    </div>

                    <div runat="server" id="HistoricoContainer" visible="false">

                        <h2>Historial de Interacciones</h2>

                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label27" runat="server" Text="Restaurante:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox19" runat="server" CssClass="textbox" Enabled="false"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label34" runat="server" Text="Pedido:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox23" runat="server" CssClass="textbox" Enabled="false"></asp:TextBox>
                                </li>
                                <ul />
                        </div>

                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label35" runat="server" Text="Repartidor:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox24" runat="server" CssClass="textbox" Enabled="false"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label36" runat="server" Text="Dias Transcurridos:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox25" runat="server" CssClass="textbox" Enabled="false"></asp:TextBox>
                                </li>
                                <ul />
                        </div>
                        <br />
                        <div class="HistoricoDiv">
                            <asp:Panel ID="Panel1" runat="server">
                            </asp:Panel>
                        </div>
                        <br />
                        <div class="form_default" runat="server" id="div1" style="margin-left: -180px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label60" runat="server" Text="Descripción:"></asp:Label>
                                </li>
                                <li>
                                    <asp:TextBox ID="TextBox49" runat="server" CssClass="textbox validate[required]" Width="450px"></asp:TextBox>
                                </li>
                            </ul>
                        </div>
                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label61" runat="server" Text="Status Queja:"></asp:Label>
                                </li>
                                <li>
                                    <asp:DropDownList ID="DropDownList13" runat="server" CssClass="textbox validate[required]">
                                        <asp:ListItem Value="">-Selecciona-</asp:ListItem>
                                        <asp:ListItem Value="1">Abierta</asp:ListItem>
                                        <asp:ListItem Value="0">Cerrada</asp:ListItem>
                                    </asp:DropDownList>
                                </li>
                            </ul>
                        </div>

                    </div>

                    <%-- ***************************************** NUEVO PEDIDO ***************************************** --%>

                    <div id="nuevo_pedido" runat="server" visible="false">

                        <asp:HiddenField ID="HiddenField4" runat="server" Value="0" />
                        <div class="form_default" style="margin-left: -50px;">
                            <ul>
                                <li>

                                    <asp:Label ID="Label38" runat="server" Text="No. Pedido VOLO:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox26" runat="server" CssClass="textbox validate[required]" AutoPostBack="true" TextMode="Number"></asp:TextBox></a></li>
                                &nbsp;
                            <asp:Image ID="Image1" runat="server" Height="20px" ImageAlign="Middle" ImageUrl="~/Images/success.png" Visible="False" Width="20px" />
                                <li>
                                    <asp:Label ID="Label66" runat="server" Text="Tienda:"></asp:Label></li>
                                <li>
                                    <asp:DropDownList ID="DropDownList15" runat="server" CssClass="textbox validate[required]" AutoPostBack="True">
                                    </asp:DropDownList></li>

                            </ul>


                        </div>
                        <div class="form_default" style="margin-left: -20px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label39" runat="server" Text="Descripción:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox27" runat="server" CssClass="textbox validate[required]" Width="350px"></asp:TextBox>
                                </li>

                                <li style="margin-left: 190px;">
                                    <asp:TextBox ID="TextBox29" runat="server" CssClass="textbox validate[required]" Width="50px" TextMode="Number"></asp:TextBox>
                                </li>


                            </ul>
                        </div>


                        <div class="form_default" style="margin-left: -180px;">
                            <ul>

                                <li>
                                    <asp:Label ID="Label40" runat="server" Text="Instrucciones especiales:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox28" runat="server" CssClass="textbox validate[required]" AutoPostBack="true" TextMode="MultiLine" Width="410px"></asp:TextBox></a></li>
                            </ul>
                        </div>

                        <asp:CheckBox ID="CheckBox2" runat="server" Text="Requiere Factura" AutoPostBack="true" />
                        <br />
                        <br />
                        <div class="form_default" style="margin-left: -50px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label63" runat="server" Text="Latitud:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox53" runat="server" CssClass="textbox validate[required,custom[number]]" AutoPostBack="true"></asp:TextBox></li>
                                <li>
                                    <asp:Label ID="Label65" runat="server" Text="Longitud:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox52" runat="server" CssClass="textbox validate[required,custom[number]]" AutoPostBack="true"></asp:TextBox></li>
                            </ul>
                        </div>




                    </div>

                    <div runat="server" id="MapContainer2" visible="false">
                        <div id="map2"></div>
                    </div>
                    <script>




                        function initMap2(dlat, dlon) {


                            var uluru = { lat: dlat, lng: dlon };
                            var map = new google.maps.Map(document.getElementById('map2'), {
                                zoom: 18,
                                center: uluru
                            });
                            var marker = new google.maps.Marker({
                                position: uluru,
                                map: map,
                                animation: google.maps.Animation.DROP,
                                draggable: false
                            });
                        }
                    </script>
                    <script
                        src="https://maps.googleapis.com/maps/api/js?key=AIzaSyCV_NilGFERUrwMoandvFuNLbFj6ymTzBI&callback=initMap2">
                    </script>


                    <%-- ***************************************** CANCELAR PEDIDO ***************************************** --%>

                    <div id="cancelarPedido" runat="server" visible="false">

                        <div class="quejas-abiertas">
                            <asp:GridView ID="GridView5" runat="server">
                            </asp:GridView>
                        </div>

                        <br />

                        <asp:HiddenField ID="HiddenField5" runat="server" Value="0" />
                        <div class="form_default">
                            <ul>
                                <li>

                                    <asp:Label ID="Label43" runat="server" Text="No. Pedido VOLO:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox30" runat="server" CssClass="textbox validate[required]" TextMode="Number"></asp:TextBox></a></li>
                                &nbsp;
                            <asp:Image ID="Image2" runat="server" Height="20px" ImageAlign="Middle" ImageUrl="~/Images/success.png" Visible="False" Width="20px" />
                            </ul>
                        </div>
                        <div class="form_default" style="margin-left: -20px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label44" runat="server" Text="Descripción:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox31" runat="server" CssClass="textbox validate[required]" AutoPostBack="true" Enabled="False" Width="350px"></asp:TextBox>
                                </li>

                                <li style="margin-left: 190px;">
                                    <asp:TextBox ID="TextBox32" runat="server" CssClass="textbox validate[required]" AutoPostBack="true" Enabled="False" Width="50px"></asp:TextBox>
                                </li>


                            </ul>
                        </div>
                    </div>

                    <%-- ***************************************** CONSULTA PEDIDO ***************************************** --%>

                    <div id="ConsultaPedidos" runat="server" visible="false">

                        <div class="quejas-abiertas">
                            <asp:GridView ID="GridView6" runat="server">
                            </asp:GridView>
                        </div>

                        <br />

                        <asp:HiddenField ID="HiddenField6" runat="server" Value="0" />
                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label45" runat="server" Text="No. Pedido VOLO:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox33" runat="server" CssClass="textbox"></asp:TextBox></a></li>
                                &nbsp;
                            <asp:Image ID="Image3" runat="server" Height="20px" ImageAlign="Middle" ImageUrl="~/Images/success.png" Visible="False" Width="20px" />
                            </ul>
                        </div>
                        <div class="form_default" style="margin-left: -20px;">
                            <ul>
                                <li>
                                    <asp:Label ID="Label46" runat="server" Text="Descripción:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox34" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False" Width="350px"></asp:TextBox>
                                </li>

                                <li style="margin-left: 190px;">
                                    <asp:TextBox ID="TextBox35" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False" Width="50px"></asp:TextBox>
                                </li>
                            </ul>
                        </div>


                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label47" runat="server" Text="Status:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox36" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label48" runat="server" Text="Repartidor:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox37" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                            </ul>
                        </div>

                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label49" runat="server" Text="Fecha Captura:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox38" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label50" runat="server" Text="Fecha Cancelacion:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox39" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                            </ul>
                        </div>

                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label51" runat="server" Text="Fecha Aceptado:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox40" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label52" runat="server" Text="Fecha Rep en Sucursal:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox41" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                            </ul>
                        </div>
                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label53" runat="server" Text="Fecha Sucursal:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox42" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label54" runat="server" Text="Fecha Recoleccion:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox43" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                            </ul>
                        </div>
                        <div class="form_default">
                            <ul>
                                <li>
                                    <asp:Label ID="Label55" runat="server" Text="Fecha Entrega:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox44" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                                <li>
                                    <asp:Label ID="Label56" runat="server" Text="DLA:"></asp:Label></li>
                                <li>
                                    <asp:TextBox ID="TextBox45" runat="server" CssClass="textbox" AutoPostBack="true" Enabled="False"></asp:TextBox>
                                </li>
                            </ul>
                        </div>
                    </div>

                    <asp:HiddenField ID="latitud" runat="server" />
                    <asp:HiddenField ID="longitud" runat="server" />

                    <div runat="server" id="mapContainer" visible="false">
                        <div id="map" style="background-image: url(../images/motito.svg); background-repeat: no-repeat; background-position: center center; background-size: 70% 80%"></div>
                    </div>
                    <script>




                        function initMap(dlat, dlon) {


                            var uluru = { lat: dlat, lng: dlon };
                            var map = new google.maps.Map(document.getElementById('map'), {
                                zoom: 18,
                                center: uluru
                            });
                            var marker = new google.maps.Marker({
                                position: uluru,
                                map: map,
                                animation: google.maps.Animation.DROP,
                                draggable: false
                            });
                        }
                    </script>
                    <script
                        src="https://maps.googleapis.com/maps/api/js?key=AIzaSyCV_NilGFERUrwMoandvFuNLbFj6ymTzBI&callback=initMap">
                    </script>

                    <a id="miEnlace" href="#modalQueja"></a>
                    <a id="miError" href="#modalError"></a>

                </ContentTemplate>
            </asp:UpdatePanel>

            <br />


            <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                <ContentTemplate>
                    &nbsp;&nbsp;<asp:CheckBox ID="CheckBox1" runat="server" Text="Nuevo Cliente" />
                </ContentTemplate>
            </asp:UpdatePanel>
            <br />
            <asp:Button ID="Button1" runat="server" Text="Guardar" CssClass="Button" Width="70px" Height="20px" />


        </div>


    </div>

</asp:Content>


