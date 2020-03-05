Imports System.Data.SqlClient
Public Class Maestra
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        LoginStatus2.LogoutText = "Cerrar Sesion (" & Request.Cookies("Usersettings")("Username") & ")"



        If Request.Cookies("Usersettings")("SU") = True Then

        ElseIf Request.Cookies("Usersettings")("Area") = 0 And Request.Cookies("Usersettings")("Puesto") = 0 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 0 And Request.Cookies("Usersettings")("Puesto") = 1 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 0 And Request.Cookies("Usersettings")("Puesto") >= 3 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 1 And Request.Cookies("Usersettings")("Puesto") = 2 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 1 And Request.Cookies("Usersettings")("Puesto") >= 3 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 2 And Request.Cookies("Usersettings")("Puesto") = 2 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 2 And Request.Cookies("Usersettings")("Puesto") >= 3 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 3 And Request.Cookies("Usersettings")("Puesto") = 2 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 3 And Request.Cookies("Usersettings")("Puesto") >= 3 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 4 And Request.Cookies("Usersettings")("Puesto") >= 3 Then

        ElseIf Request.Cookies("Usersettings")("Area") = 5 And Request.Cookies("Usersettings")("Puesto") >= 2 Then

        End If


    End Sub



End Class