Imports System.Data.SqlClient

Public Class WebForm1
    Inherits System.Web.UI.Page

    Public Function GetPassStatus(UserName As String) As Boolean

        Dim sql As String = "SELECT COUNT(*) FROM SYS_empleados WHERE id_ccs = @user AND (status = 2 OR status = 7) AND su IS NOT NULL"
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("db").ToString())

            conn.Open()
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@user", UserName)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            If count = 0 Then
                Return False
            Else
                Return True

            End If

        End Using



    End Function


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If GetPassStatus(Request.Cookies("UserSettings")("Username")) = False Then
            Response.Redirect("Password.aspx")
        Else

        End If




    End Sub

End Class