Imports System.Data.SqlClient
Public Class Password
    Inherits System.Web.UI.Page

    Dim Alerta As New Alertas
    Dim msgtipo(20) As Integer
    Dim msgmensaje(20) As String

    Sub Cambiar()

        Dim x As New Funciones

        If TextBox1.Text = TextBox2.Text Then

            Dim strConnString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
            Dim strQuery As String = "UPDATE SYS_empleados SET pass_ccs = '" & x.passcrypt(TextBox1.Text) & "', su = '0' WHERE id_ccs = '" & Request.Cookies("UserSettings")("Username") & "'"
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

        Else

            msgtipo(0) = 4
            msgmensaje(0) = "¡Los passwords introducidos no coinciden!"
            Alerta.NewShowAlert(msgtipo, msgmensaje, Me)

        End If

    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cambiar()
    End Sub

End Class