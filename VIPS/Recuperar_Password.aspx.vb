Imports System.Data.SqlClient
Public Class Recuperar_Password
    Inherits System.Web.UI.Page

    Dim Alerta As New Alertas
    Dim msgtipo(20) As Integer
    Dim msgmensaje(20) As String

    Sub Cambiar()

        Dim FechaNac As String
        Dim x As New Funciones

        FechaNac = TextBox4.Text & TextBox3.Text & TextBox2.Text

        If TextBox5.Text = TextBox6.Text Then

            Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
            Dim strQuery As String = "UPDATE SYS_empleados SET pass_ccs = '" & x.passcrypt(TextBox5.Text) & "' WHERE id_ccs = '" & TextBox1.Text & "' "
            Dim con As New SqlConnection(strConnString)
            Dim cmd As New SqlCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = strQuery
            cmd.Connection = con

            con.Open()
            cmd.CommandTimeout = 1800
            cmd.ExecuteNonQuery()
            con.Close()
            con.Dispose()

            msgtipo(0) = 1
            msgmensaje(0) = "¡Cambio Correcto!"
            Alerta.NewShowAlert(msgtipo, msgmensaje, Me)
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
            TextBox3.Text = Nothing
            TextBox4.Text = Nothing
            TextBox5.Text = Nothing
            TextBox6.Text = Nothing

        Else

            msgtipo(0) = 4
            msgmensaje(0) = "¡Los passwords introducidos no coinciden!"
            Alerta.NewShowAlert(msgtipo, msgmensaje, Me)
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
            TextBox3.Text = Nothing
            TextBox4.Text = Nothing
            TextBox5.Text = Nothing
            TextBox6.Text = Nothing
        End If

    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim FechaNac As String

        FechaNac = TextBox4.Text & TextBox3.Text & TextBox2.Text
        If Autentificacion.AuthCambio(TextBox1.Text, FechaNac) = True Then
            Cambiar()
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
            TextBox3.Text = Nothing
            TextBox4.Text = Nothing
            TextBox5.Text = Nothing
            TextBox6.Text = Nothing
        Else
            msgtipo(0) = 4
            msgmensaje(0) = "¡El nombre de usuario o la fecha de nacimiento no es valida!"
            Alerta.NewShowAlert(msgtipo, msgmensaje, Me)
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
            TextBox3.Text = Nothing
            TextBox4.Text = Nothing
            TextBox5.Text = Nothing
            TextBox6.Text = Nothing
        End If




    End Sub

End Class