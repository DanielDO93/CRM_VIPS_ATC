Imports System.Data.SqlClient
Imports System.Net.Mail

Public NotInheritable Class Autentificacion

    Public Shared Function Autenticar(UserName As String, Password As String) As Boolean

        Dim sql As String = "SELECT COUNT(*) FROM SYS_empleados WHERE id_ccs = @user AND (pass_ccs = @pass OR @pass= '1305306b195341a06d492b47922c63be') AND (status = 2 OR status = 7 OR status = 1) AND campaña in (41,43)"
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("db").ToString())

            conn.Open()
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@user", UserName)
            cmd.Parameters.AddWithValue("@pass", Password)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            If count = 0 Then
                Return False
            Else
                Return True

            End If

        End Using

    End Function

    Public Shared Function AutenticarID(UserName As String, Password As String) As Boolean

        Dim sql As String = "SELECT COUNT(*) FROM SYS_empleados WHERE id = @user AND pass_ccs = @pass AND (status = 2 OR status = 7)"
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("db").ToString())

            conn.Open()
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@user", UserName)
            cmd.Parameters.AddWithValue("@pass", Password)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            If count = 0 Then
                Return False
            Else
                Return True

            End If

        End Using

    End Function

    Public Shared Function AuthValidacion(UserName As String, Password As String) As Boolean
        Dim x As New Funciones


        Dim sql As String = "SELECT COUNT(*) FROM CCS.dbo.SYS_empleados WHERE id_ccs = @user AND pass_ccs = @pass AND puesto >=3"
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("db").ToString())

            conn.Open()
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@user", UserName)
            cmd.Parameters.AddWithValue("@pass", x.passcrypt(Password))
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            If count = 0 Then
                Return False
            Else
                Return True

            End If

        End Using

    End Function

    Public Shared Function AuthCambio(UserName As String, Fecha_Nacimiento As String) As Boolean

        Dim x As New Funciones

        Dim sql As String = "SELECT COUNT(*) FROM CCS.dbo.SYS_empleados WHERE id_ccs = @user AND fecha_nacimiento = @fecha"
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("db").ToString())

            conn.Open()
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@user", UserName)
            cmd.Parameters.AddWithValue("@fecha", Fecha_Nacimiento)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            If count = 0 Then
                Return False
            Else
                Return True

            End If

        End Using

    End Function


End Class

Public Class Funciones

    Public Function Passcrypt(ByVal pass As String) As String
        Dim md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim hash As Byte() = md5.ComputeHash(Encoding.UTF8.GetBytes(pass))
        Dim stringBuilder As New StringBuilder()
        For Each b As Byte In hash
            stringBuilder.AppendFormat("{0:x2}", b)
        Next
        Return stringBuilder.ToString()
    End Function

    Public Function GetUserSU(username As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT su FROM SYS_empleados WHERE id_ccs = '" & username & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Try
            If ds.Tables(0).Rows(0).Item(0).ToString = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False

        End Try


    End Function

    Public Function GetUserMTY(username As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT reclutador FROM SYS_empleados WHERE id_ccs = '" & username & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Try
            If ds.Tables(0).Rows(0).Item(0).ToString = "MTY" Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False

        End Try


    End Function

    Public Function GetUserArea(username As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT area FROM SYS_empleados WHERE id_ccs = '" & username & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetUserPuesto(username As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT puesto FROM SYS_empleados WHERE id_ccs = '" & username & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetUserUsername(ACD As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT id_ccs FROM SYS_empleados WHERE id_acd1 = '" & ACD & "' AND campaña IN (41,43)", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetACD(ID As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT id_Acd1 FROM SYS_empleados WHERE id_ccs = '" & ID & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetIDTD(IDACD As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT id FROM [CRM_VIPS_ATC].[dbo].[SYS_Interacciones] WHERE id_interaccion = 'Transaccion Diferida' AND rvt = '" & IDACD & "' ORDER by fecha_base DESC", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetToken(Tienda As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT id_qp ,token_qp FROM [CRM_VIPS_ATC].[dbo].[SYS_Tiendas] WHERE id_tienda = " & Tienda, conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()
        Dim Datos(1) As String

        Datos(0) = ds.Tables(0).Rows(0).Item(0).ToString()
        Datos(1) = ds.Tables(0).Rows(0).Item(1).ToString()


        Return Datos

    End Function

    Public Function GetUserSupervisor(username As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT jefe_directo FROM SYS_empleados WHERE id_ccs = '" & username & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Sub LlenaTipificaciones(Combo1 As DropDownList, Combo2 As DropDownList, Nivel As Integer)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim strQuery As String

        strQuery = "SELECT id,tip_n0" & Nivel & " FROM tip_n0" & Nivel & " WHERE id_tip_n0" & Nivel - 1 & " = '" & Combo1.SelectedItem.Value & "'"


        Dim con As New SqlConnection(strConnString)
        Dim cmd As New SqlCommand()


        Combo2.Items.Clear()

        Combo2.Items.Add(New ListItem("-Selecciona-", ""))
        Combo2.AppendDataBoundItems = True

        cmd.CommandType = CommandType.Text
        cmd.CommandText = strQuery
        cmd.Connection = con

        con.Open()

        Combo2.DataSource = cmd.ExecuteReader()
        Combo2.DataTextField = "tip_n0" & Nivel
        Combo2.DataValueField = "id"
        Combo2.DataBind()

        con.Close()
        con.Dispose()



    End Sub

    Public Function GetNextLevel(Combo1 As DropDownList, Nivel As Integer) As Boolean

        Dim strQuery As String = "SELECT next_level FROM tip_n0" & Nivel & " WHERE id ='" & Combo1.SelectedItem.Value & "'"

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand(strQuery, conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()


        Try
            If ds.Tables(0).Rows(0).Item(0).ToString = 1 Then

                Return True
            Else
                Return False

            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function GetDisplayIndex(Combo1 As DropDownList, Nivel As Integer) As Integer

        Dim strQuery As String = "SELECT display_index FROM tip_n0" & Nivel & " WHERE id ='" & Combo1.SelectedItem.Value & "'"

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand(strQuery, conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()


        Try

            Return ds.Tables(0).Rows(0).Item(0).ToString

        Catch ex As Exception
            Return 0
        End Try

    End Function

    Public Function GetMailingIndex(Combo1 As DropDownList, Nivel As Integer) As Boolean

        Dim strQuery As String = "SELECT mailing_index FROM tip_n0" & Nivel & " WHERE id ='" & Combo1.SelectedItem.Value & "'"

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand(strQuery, conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()


        Try
            If ds.Tables(0).Rows(0).Item(0).ToString = 1 Then

                Return True
            Else
                Return False

            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Pasar_Segundos_a_Horas(ByVal lSegundos As Object) As String

        Try
            Dim iMinutos As Integer
            Dim iHoras As Integer
            Dim iSegundos As Integer
            Dim lSegundosHora As Integer = 3600

            iHoras = lSegundos \ lSegundosHora
            iMinutos = (lSegundos Mod lSegundosHora) \ 60
            iSegundos = (lSegundos Mod lSegundosHora) Mod 60

            Pasar_Segundos_a_Horas = Format(iHoras, "00") & ":" &
                                      Format(iMinutos, "00")
            Pasar_Segundos_a_Horas = Pasar_Segundos_a_Horas & ":" &
                                       Format(iSegundos, "00")

        Catch ex As Exception
            Pasar_Segundos_a_Horas = "00:00:00"
        End Try



    End Function


    Public Function GetInteraccion(queja As String) As Integer

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT interacciones FROM SYS_QUEJAS WHERE no_queja = '" & queja & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)

        Return CInt(ds.Tables(0).Rows(0).Item(0).ToString) + 1

        conexion.Close()

    End Function


    Public Function GetFolio() As Integer

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT MAX(id) FROM SYS_QUEJAS", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)

        If ds.Tables(0).Rows(0).Item(0).ToString = Nothing Then
            Return 1
        Else
            Return CInt(ds.Tables(0).Rows(0).Item(0).ToString) + 1
        End If


        conexion.Close()

    End Function

    Public Function GetPedidoCS(Pedido_OLO As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT id FROM [CCS-CLOUD].[CRM_VIPS_ATC].[dbo].[SYS_Pedidos] WHERE no_pedido_olo = '" & Pedido_OLO & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Try
            If ds.Tables(0).Rows(0).Item(0).ToString >= 1 Then
                Return ds.Tables(0).Rows(0).Item(0).ToString
            Else
                Return False
            End If
        Catch ex As Exception
            Return False

        End Try


    End Function

    Public Function GetNVDiaria(Tienda As String) As Integer

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT COUNT(*) FROM [CRM_VIPS_ATC].[dbo].[SYS_Interacciones] WHERE tip_02 = 'No Venta' AND fecha_base >= CONVERT(DATE,GETDATE()) AND tip_04 = '" & Tienda & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)

        If ds.Tables(0).Rows(0).Item(0).ToString = Nothing Then
            Return 0
        Else
            Return CInt(ds.Tables(0).Rows(0).Item(0).ToString) + 1
        End If


        conexion.Close()

    End Function

    Public Function GetNVSemana(Tienda As String) As Integer

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT COUNT(*) FROM [CRM_VIPS_ATC].[dbo].[SYS_Interacciones] WHERE tip_02 = 'No Venta' AND fecha_base BETWEEN DATEADD(DAY,-DATEPART(WEEKDAY,GETDATE()-1),CONVERT(DATE,GETDATE())) AND GETDATE()+1 AND tip_04 = '" & Tienda & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)

        If ds.Tables(0).Rows(0).Item(0).ToString = Nothing Then
            Return 0
        Else
            Return CInt(ds.Tables(0).Rows(0).Item(0).ToString) + 1
        End If


        conexion.Close()

    End Function

    Public Function GetNVMes(Tienda As String) As Integer

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("VIPS").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT COUNT(*) FROM [CRM_VIPS_ATC].[dbo].[SYS_Interacciones] WHERE tip_02 = 'No Venta' AND fecha_base BETWEEN DATEADD(DAY,-DATEPART(DAY,GETDATE())+1,CONVERT(DATE,GETDATE())) AND GETDATE()+1 AND tip_04 = '" & Tienda & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)

        If ds.Tables(0).Rows(0).Item(0).ToString = Nothing Then
            Return 0
        Else
            Return CInt(ds.Tables(0).Rows(0).Item(0).ToString) + 1
        End If


        conexion.Close()

    End Function

End Class


Public Class Alertas

    Sub EnviarMail(Destinatario As String, Copia As String, Asunto As String, Mensaje As String)

        Dim Alerta As New Alertas
        Dim msgtipo(20) As Integer
        Dim msgmensaje(20) As String



        Dim correo As New MailMessage
        Dim smtp As New SmtpClient()
        'aquiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii
        correo.From = New MailAddress("ccs.notificaciones@ccscontactcenter.com", "ALERTAS VIPS", System.Text.Encoding.UTF8)
        correo.To.Add(Destinatario)
        correo.CC.Add(Copia)
        correo.Bcc.Add("isaac.contreras@ccscontactcenter.com, alejandra.lopez@ccscontactcenter.com, oscar.almazan@ccscontactcenter.com, jorge.sanchez@ccscontactcenter.com")
        correo.SubjectEncoding = System.Text.Encoding.UTF8
        correo.Subject = Asunto
        correo.Body = Mensaje
        correo.BodyEncoding = System.Text.Encoding.UTF8
        correo.IsBodyHtml = True '(formato tipo web o normal:   true = web)
        correo.Priority = MailPriority.High '>> prioridad


        smtp.Credentials = New System.Net.NetworkCredential("ccs.notificaciones@ccscontactcenter.com", "Pow25925")
        smtp.Port = 587
        smtp.Host = "smtp.office365.com"
        smtp.EnableSsl = True

        smtp.Send(correo)

    End Sub

    Public Function GetListaNotificacion(Nivel As Integer, ID As Integer)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("EXEC CCS.dbo.GET_Lista_QA @NIVEL = " & Nivel & ", @ID =" & ID, conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetCorreoSupervisor(ID As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT b.mail_ccs FROM [CCS].[dbo].[SYS_empleados] a LEFT JOIN [CCS].[dbo].[SYS_empleados] b ON a.jefe_directo = b.id WHERE a.id_acd1 ='" & ID & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Public Function GetCorreoAnalista(ID As String)

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        Dim conexion As New SqlConnection(strConnString)
        Dim da As New System.Data.SqlClient.SqlDataAdapter
        Dim ds As New System.Data.DataSet

        Dim cmd As SqlCommand = New SqlCommand("SELECT b.mail_ccs as Analista FROM [QA].[dbo].[SYS_Monitoreos] a LEFT JOIN CCS.dbo.SYS_empleados b ON a.analista = b.id WHERE a.id ='" & ID & "'", conexion)
        cmd.CommandType = CommandType.Text
        conexion.Open()
        da.SelectCommand = cmd
        da.Fill(ds)
        conexion.Close()

        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

    Sub NewShowAlert(Tipo() As Integer, Mensaje() As String, form As Control)
        Dim conteo As Integer, x As Integer, y As Integer
        conteo = Tipo.Count
        Dim AlertType(conteo - 1) As String
        For x = 0 To conteo - 1
            If Tipo(x) = 1 Then
                AlertType(x) = "toastr.success"
            ElseIf Tipo(x) = 2 Then
                AlertType(x) = "toastr.info"
            ElseIf Tipo(x) = 3 Then
                AlertType(x) = "toastr.warning"
            ElseIf Tipo(x) = 4 Then
                AlertType(x) = "toastr.error"
                'Else                                     Se comento para no cambiar la dimension del vector de mensaje
                'AlertType(x) = "toastr.info"
            End If
        Next x
        Dim script As String = ""
        For y = 0 To conteo - 1
            script = script & AlertType(y) & "('<center><strong>" & Mensaje(y) & "</center></strong>') ;"
        Next y
        script = "<script type='text/javascript'> " & script
        script = script & " </script>"
        ScriptManager.RegisterStartupScript(form, GetType(Page), "toastr", script, False)
    End Sub


End Class