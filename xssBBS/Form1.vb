<Security.Permissions.PermissionSet(Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class Form1
    Private Const myURL As String = "http://example.jp/bbs.cgi"
    <Serializable> Private Class clsEntry
        Public Property Author As String = ""
        Public Property EMail As String = ""
        Public Property [Date] As String = ""
        Public Property Title As String = ""
        Public Property Body As String = ""
        Public Property [Event] As String = ""
    End Class
    Private colEntries As New Queue(Of clsEntry)
    Private errormsg As String = ""
    Private Const strPassword As String = "RealKusomonIsHere!!!"

    ''' <summary>
    ''' Form Load event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Call initColEntries()
            WebBrowser1.Navigate("about:blank")
            WebBrowser1.ObjectForScripting = Me
            Call writeForm()
            Me.Visible = True

            Dim ie As New SHDocVw.InternetExplorer
            ie.Visible = False
            ie.Silent = True
            ie.Navigate2("about:blank")

            If My.Application.CommandLineArgs.Count <= 0 Then
                MsgBox("This program will launch or kill ""Internet Explorer"" for any time.", MsgBoxStyle.Information, "Information")
                Call killIEs()
            Else
                'Try
                '    Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                '    Dim c As clsEntry = bf.Deserialize(New System.IO.MemoryStream(System.Convert.FromBase64String(My.Application.CommandLineArgs(0))))
                '    colEntries.Enqueue(c)
                '    Call writeForm()
                '    Call killIEs()
                '    Dim th As New launchIE_dg(AddressOf launchIE)
                '    th.BeginInvoke(c.Event, c, Nothing, Nothing)
                'Catch ex As Exception
                'End Try
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Make first BBS entry
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initColEntries()
        Try
            colEntries.Clear()

            Dim c As New clsEntry
            c.Author = "admin"
            c.Body = "Hi, guys. Feel free to message me. I'll make a regular check."
            c.Date = Now.ToString("ddd, dd MMM yyyy HH:mm:ss", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            c.EMail = "ymzkei5@example.jp"
            c.Title = "I'm admin."
            colEntries.Enqueue(c)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Kill "iexplorer"
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub killIEs()
        Try
            For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcesses
                If p.ProcessName.ToLower = "iexplore" Then
                    p.Kill()
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Show BBS entries
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub writeForm()
        Try
            WebBrowser1.Document.OpenNew(False)
            WebBrowser1.Document.Write(getHTML(False))
            WebBrowser1.Document.GetElementsByTagName("form")(0).AttachEventHandler("onsubmit", AddressOf onSubmitAddEntryButton)
            WebBrowser1.Document.GetElementsByTagName("form")(1).AttachEventHandler("onsubmit", AddressOf onSubmitLoginButton)
            errormsg = ""
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Make HTML
    ''' </summary>
    ''' <param name="isAdmin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getHTML(isAdmin As Boolean) As String
        Try
            Dim html As New System.IO.StringWriter
            For Each c As clsEntry In colEntries.ToArray.Reverse
                Try
                    html.WriteLine("<hr><span style='font-weight:bold;background-color:#f4e242'>" + c.Title + "</span> by <a href='mailto:" + c.EMail + "'>" + c.Author + "</a> - " + c.Date + "<hr><ul>" + c.Body + "</ul>")
                Catch ex As Exception
                End Try
            Next
            Return "<html><body><h2>Old and Legacy ""Hakoniwa"" BBS</h2>" + _
                   "<form action='" + myURL + "' method='get'>" + errormsg + "<table border='0'>" + _
                   "<tr><td>Name:</td><td><input type='text' name='author' value='' size='50'></td></tr>" + _
                   "<tr><td nowrap>Mail:</td><td><input type='email' name='email' value='' size='50'></td></tr>" + _
                   "<tr><td>Title:</td><td><input type='text' name='title' value='' size='70'></td></tr>" + _
                   "<tr><td>Comments:</td><td><textarea name='body' cols='70' rows='5'></textarea></td></tr>" + _
                   "<tr><td colspan='2'><input type='submit' value='Add Entry'></td></tr>" + _
                   "</table></form><form action='" + myURL + "' method='get'><hr>Admin Password: <input type='password' name='pass' id='pass' value='" + IIf(isAdmin, strPassword, "") + "'> " + _
                   "<input type='submit' value='LogIn' id='admlogin'><hr></form><p>" + vbCrLf + html.ToString + "</body></html>"
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Login Button clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub onSubmitLoginButton(sender As Object, e As EventArgs)
        Try
            Dim pass As String = WebBrowser1.Document.GetElementById("pass").GetAttribute("value")
            If pass = strPassword Then
                WebBrowser1.Document.OpenNew(False)
                WebBrowser1.Document.Write("FLAG{CAN_YOU_ENJOY_ENJYO?}")
            Else
                errormsg = "<hr color='red'><font color='red'>Wrong admin password!</font><hr color='red'>"
                Call writeForm()
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Add Entry Button clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub onSubmitAddEntryButton(sender As Object, e As EventArgs)
        Try
            Dim author As String = WebBrowser1.Document.GetElementById("author").GetAttribute("value")
            Dim title As String = WebBrowser1.Document.GetElementById("title").GetAttribute("value")
            Dim email As String = WebBrowser1.Document.GetElementById("email").GetAttribute("value")
            Dim body As String = WebBrowser1.Document.GetElementById("body").InnerText
            errormsg = writeBBS(author, title, email, body)
            Call writeForm()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Write BBS entry
    ''' </summary>
    ''' <param name="author"></param>
    ''' <param name="title"></param>
    ''' <param name="email"></param>
    ''' <param name="body"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function writeBBS(author As String, title As String, email As String, body As String) As String
        Try
            Dim ret As String = ""
            If author = "" OrElse title = "" OrElse email = "" OrElse body = "" Then
                ret = "Fill in all items."
            End If
            If ret = "" Then
                Dim c As New clsEntry
                c.Author = System.Web.HttpUtility.HtmlEncode(author)
                c.Title = System.Web.HttpUtility.HtmlEncode(title)
                c.EMail = System.Web.HttpUtility.HtmlEncode(email).Replace("&quot;", """")
                c.Body = System.Web.HttpUtility.HtmlEncode(body)
                c.Date = Now.ToString("ddd, dd MMM yyyy HH:mm:ss", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"))
                colEntries.Enqueue(c)
                Dim evnt As String = System.Text.RegularExpressions.Regex.Match(c.EMail, "'(?:.*[^a-z])?(on[a-z]+)=", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups(1).Value
                If evnt <> "" Then
                    Try
                        c.Event = evnt
                        Call killIEs()
                        Dim th As New launchIE_dg(AddressOf launchIE)
                        th.BeginInvoke(evnt, c, Nothing, Nothing)
                    Catch ex As Exception
                    End Try
                End If
            Else
                ret = "<hr color='red'><font color='red'>" + ret + "</font><hr color='red'>"
            End If
            Return ret
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Launch Admin's Internet Explorer
    ''' </summary>
    ''' <param name="evnt"></param>
    ''' <param name="cls"></param>
    ''' <remarks></remarks>
    Private Delegate Sub launchIE_dg(evnt As String, cls As clsEntry)
    Private Sub launchIE(evnt As String, cls As clsEntry)
        Try
            Dim ie As New SHDocVw.InternetExplorer
            Try
                ie.Visible = False
                ie.Silent = True
                ie.Navigate2("about:blank")
                For i As Integer = 1 To 10
                    If ie.ReadyState = SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE OrElse ie.ReadyState = SHDocVw.tagREADYSTATE.READYSTATE_INTERACTIVE Then Exit For
                    System.Threading.Thread.Sleep(500)
                Next
                AddHandler ie.BeforeNavigate2, AddressOf ieBeforeNavigate2
                DirectCast(ie.Document, mshtml.IHTMLDocument2).Open()
                DirectCast(ie.Document, mshtml.IHTMLDocument2).write(getHTML(True))
                System.Threading.Thread.Sleep(2000)
                DirectCast(ie.Document, mshtml.IHTMLDocument3).getElementsByTagName("A")(0).fireEvent(evnt)
                System.Threading.Thread.Sleep(2000)
                DirectCast(ie.Document, mshtml.IHTMLDocument3).getElementsByTagName("FORM")(1).Submit()
            Finally
                Try
                    ie.Quit()
                Catch ex As Exception
                End Try
            End Try
        Catch ex As Exception
            'If ex.Message.Contains("0x800A01B6") Then
            '    Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            '    Dim ms As New System.IO.MemoryStream
            '    bf.Serialize(ms, cls)
            '    Debug.Print(System.Convert.ToBase64String(ms.ToArray))
            '    System.Diagnostics.Process.Start(System.Windows.Forms.Application.ExecutablePath, System.Convert.ToBase64String(ms.ToArray))
            '    End
            'End If
        End Try
    End Sub

    ''' <summary>
    ''' Query String parsing
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function writeBBSbyURL(url As String) As String
        Try
            Dim author As String = System.Web.HttpUtility.UrlDecode(System.Text.RegularExpressions.Regex.Match(url.ToString, "[?&]author=([^&]*)").Groups(1).Value)
            Dim title As String = System.Web.HttpUtility.UrlDecode(System.Text.RegularExpressions.Regex.Match(url.ToString, "[?&]title=([^&]*)").Groups(1).Value)
            Dim email As String = System.Web.HttpUtility.UrlDecode(System.Text.RegularExpressions.Regex.Match(url.ToString, "[?&]email=([^&]*)").Groups(1).Value)
            Dim body As String = System.Web.HttpUtility.UrlDecode(System.Text.RegularExpressions.Regex.Match(url.ToString, "[?&]body=([^&]*)").Groups(1).Value)
            Return writeBBS(author, title, email, body)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Show BBS on admin's browser
    ''' </summary>
    ''' <param name="pDisp"></param>
    ''' <param name="URL"></param>
    ''' <param name="Flags"></param>
    ''' <param name="TargetFrameName"></param>
    ''' <param name="PostData"></param>
    ''' <param name="Headers"></param>
    ''' <param name="Cancel"></param>
    ''' <remarks></remarks>
    Public Sub ieBeforeNavigate2(pDisp As Object, ByRef URL As Object, ByRef Flags As Object, ByRef TargetFrameName As Object, ByRef PostData As Object, ByRef Headers As Object, ByRef Cancel As Boolean)
        Try
            If URL.ToString = "about:blank" Then
                '
            ElseIf URL.ToString.StartsWith(myURL) OrElse URL.ToString.StartsWith("about:blank?") Then
                Call writeBBSbyURL(URL)
                Cancel = True
            ElseIf URL.Host = New Uri(myURL).Host OrElse URL.ToString.StartsWith("about:") Then
                DirectCast(pDisp.Document, mshtml.IHTMLDocument2).open()
                DirectCast(pDisp.Document, mshtml.IHTMLDocument2).write("<h2>404 Not Found</h2>")
                Cancel = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Show BBS on user's browser
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub WebBrowser1_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles WebBrowser1.Navigating
        Try
            If e.Url.ToString = "about:blank" Then
                '
            ElseIf e.Url.ToString.StartsWith(myURL) OrElse e.Url.ToString.StartsWith("about:blank?") Then
                errormsg = writeBBSbyURL(e.Url.ToString)
                e.Cancel = True
                Call writeForm()
            ElseIf e.Url.Host = New Uri(myURL).Host OrElse e.Url.ToString.StartsWith("about:") Then
                WebBrowser1.Document.OpenNew(False)
                WebBrowser1.Document.Write("<h2>404 Not Found</h2>")
                e.Cancel = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Show webbrowser status message
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub WebBrowser1_StatusTextChanged(sender As Object, e As EventArgs) Handles WebBrowser1.StatusTextChanged
        Try
            ToolStripStatusLabel1.Text = WebBrowser1.StatusText
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Refresh view
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lblRefresh_Click(sender As Object, e As EventArgs) Handles lblRefresh.Click
        Try
            Call writeForm()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Clear data (delete all entries)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lblClear_Click(sender As Object, e As EventArgs) Handles lblClear.Click
        Try
            Call initColEntries()
            Call writeForm()
        Catch ex As Exception
        End Try
    End Sub
End Class
