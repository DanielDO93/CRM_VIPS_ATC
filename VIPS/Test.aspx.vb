Imports System.IO
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
Public Class Test
    Inherits System.Web.UI.Page
    Protected Sub Capture(sender As Object, e As EventArgs)
        Dim url As String = txtUrl.Text.Trim()

        Dim thread As New Thread(Sub()
                                     Using browser As New WebBrowser()
                                         browser.ScrollBarsEnabled = False
                                         browser.AllowNavigation = True
                                         browser.Navigate(url)
                                         browser.Width = 1024
                                         browser.Height = 768
                                         'AddHandler browser.DocumentCompleted, New WebBrowserDocumentCompletedEventHandler(AddressOf DocumentCompleted)
                                         'While browser.ReadyState <> WebBrowserReadyState.Complete
                                         'System.Windows.Forms.Application.DoEvents()
                                         'End While

                                         browser.Show()
                                     End Using

                                 End Sub)
        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()
        thread.Join()

        'Dim ctrl As HtmlControl = DirectCast(UpdatePanel1.FindControl("mainiFrame"), HtmlControl)
        'ctrl.Attributes("src") = "http://www.google.com"


    End Sub

    Private Sub DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        Dim browser As WebBrowser = TryCast(sender, WebBrowser)
        Using bitmap As New Bitmap(browser.Width, browser.Height)
            browser.DrawToBitmap(bitmap, New Rectangle(0, 0, browser.Width, browser.Height))
            Using stream As New MemoryStream()
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                Dim bytes As Byte() = stream.ToArray()
                imgScreenShot.Visible = True
                imgScreenShot.ImageUrl = "data:image/png;base64," + Convert.ToBase64String(bytes)
            End Using
        End Using
    End Sub

End Class