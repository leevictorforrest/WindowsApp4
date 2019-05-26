Imports System.IO
Imports System.Net

Public Class Form1
    Public linkreceived As Boolean
    Public x1, x2, y1, y2 As Integer
    Public cnt1, cnt2 As Integer
    Public ctadd As Long
    Public screencords(1024) As Integer
    Public sx1(1024) As Integer
    Public sx2(1024) As Integer
    Public sy1(1024) As Integer
    Public sy2(1024) As Integer
    Public hyplink(1024) As String
    Public msx, msy As Long


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        GetAlbumInfo()

        ' "C:\vb\vupics\Ftorrents100 Greatest Classic Rock Songs (2019)001.  Van Halen  -  Jump (2015 Remaster).mp3\1.jpg"

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        For x = 0 To ctadd
            If sx1(x) <= msx And sy1(x) <= msy And sx2(x) >= msx And sy2(x) >= msy Then
                MsgBox(Str(x))
            End If
        Next

    End Sub

    Public Sub dispic()
        Dim _picCollection As New PictureBox.ControlAccessibleObject(Me)
        Dim _picCollection1 As New PictureBox.ControlAccessibleObject(Me)



        Me.Refresh()
        For Each pic1 In My.Computer.FileSystem.GetFiles("C:\vb\vupics\Ftorrents100 Greatest Classic Rock Songs (2019)001.  Van Halen  -  Jump (2015 Remaster).mp3", FileIO.SearchOption.SearchTopLevelOnly, "*.jpg")
            Dim pp As String = Str(x1)
            Dim pic As New PictureBox
            pic.ImageLocation = pic1
            Dim rnd As New Random
            pic.Location = New Point(x1, y1)
            pic.Size = New Size(pic.Width, pic.Height)
            pic.BackColor = Color.White
            pic.Visible = True
            pic.BringToFront()



            pic.SizeMode = PictureBoxSizeMode.Zoom
            Me.Controls.Add(pic)
            x1 = x1 + pic.Width
            If x1 > Me.Width - pic.Width Then
                x1 = 1
                y1 = y1 + pic.Height
            End If

        Next
    End Sub
    Private Sub GetAlbumInfo()
        For xx = 1 To ctadd
            For Each cControl As Control In Me.Controls



                If cControl.Name = "" Then
                    Me.Controls.Remove(cControl)
                    screencords(xx) = 0
                    sx1(xx) = 0
                    sx2(xx) = 0
                    sy1(xx) = 0
                    sy2(xx) = 0
                End If
                'MessageBox.Show(cControl.Name)




                'You can put here the validation if the user have rights to access the control and enable/disable it' 




            Next
        Next
        ctadd = 0
        x1 = 10 : y1 = 10 : cnt1 = 10 : cnt2 = 10
        Dim zArtist As String = TextBox1.Text
        Dim zTrackTitle As String = TextBox2.Text
        Dim zAlbum As String = TextBox1.Text
        'MsgBox(zArtist + vbLf + zTrackTitle + vbLf + zAlbum)
        Try
            'no base exists

            Dim ss As String = "&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjMyLvVv6XiAhV1sXEKHV1LCD8Q_AUIDygC&biw=1202&bih=493"
            Dim bak As String = "&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjG0s6j6qDiAhXfThUIHVXBCTUQ_AUIDigB&biw=1280&bih=551"

            Dim mainUrl As String = "https://www.google.co.uk/search?q=" + zArtist + "+" + zTrackTitle + Me.Width.ToString + "x" + Me.Height.ToString + ss
            'https://www.google.co.uk/search?q=michael+jackson+cover+art&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjG0s6j6qDiAhXfThUIHVXBCTUQ_AUIDigB&biw=1280&bih=551
            'https://www.google.co.uk/search?q=michael+jackson+cover+art&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjG0s6j6qDiAhXfThUIHVXBCTUQ_AUIDigB&biw=1280&bih=551
            'https://www.google.co.uk/search?q=michael+jackson+beat+it&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjMyLvVv6XiAhV1sXEKHV1LCD8Q_AUIDygC&biw=1202&bih=493

            'Dim doc As HtmlAgilityPack.HtmlDocument
            Dim doc = New HtmlAgilityPack.HtmlDocument()
            Dim nod As HtmlAgilityPack.HtmlNode
            Dim wc As New WebClient()
            Dim sourcestring As String
            Dim request = CType(WebRequest.Create(mainUrl), HttpWebRequest)
            'Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; AS; rv:11.0)

            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705;)"
            '"Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705;)
            '"•Links (0.96; CYGWIN_NT-5.0 1.3.22(0.78/3/2) i686)
            request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip,deflate")
            request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
            'Dim responseText As String
            ' Dim response = request.GetResponse()

            Try
                Using response As WebResponse = request.GetResponse()
                    Using responsestream As Stream = response.GetResponseStream()
                        Using reader As StreamReader = New StreamReader(responsestream)
                            Dim webresponse As String
                            webresponse = reader.ReadToEnd()
                            sourcestring = webresponse
                            Application.DoEvents()
                        End Using
                    End Using
                End Using
            Catch ex As Exception

            End Try
            'Using response As WebResponse = request.GetResponseasync()


            ' MsgBox(sourcestring)

            ' does not unpack the contents then if source string = gzip

            'Dim sourceString1 As String = New System.Net.WebClient().DownloadString(mainUrl)
            'MsgBox(sourceString1)

            '= wc.DownloadString(mainUrl)
            Dim sourcestrin1 As String = sourcestring
            If sourcestring = "" Then GoTo thisjmp
            doc.LoadHtml(sourcestring)
            Dim st As String = TextBox1.Text
            Dim newString As String = st.Replace(vbCr, "").Replace(vbLf, "")
            st = newString
            Dim pth As String = st
            Dim s() As String = Split(pth, "\")
            Dim n As Long

            pth = ""
            For Each itm In s
                pth = pth + itm
            Next
            Dim s1() As String = Split(pth, ":")
            pth = ""
            For Each colon In s1
                pth = pth + colon
            Next
            pth = pth + TextBox1.Text + TextBox2.Text
            Dim cct As Long
            If Not IO.Directory.Exists("c:\vb\vupics\" + pth) Then
                IO.Directory.CreateDirectory("c:\vb\vupics\" + pth)
            End If

            For Each itm In IO.Directory.GetFiles("c:\vb\vupics\" + pth + "\", "*.*", searchOption:=SearchOption.TopDirectoryOnly)
                cct = cct + 1
            Next
            If cct > 1 Then
                'GoTo jmp2
            End If
            Dim ccnt As Long

scanimages:
            Try
                ccnt = ccnt + 1
                If ccnt > 2 Then
                    GoTo exxxxt
                End If
                For Each link As HtmlAgilityPack.HtmlNode In doc.DocumentNode.SelectNodes("//img[@src]")
                    linkreceived = True
                    Application.DoEvents()

                    Dim linkAddress = GetAbsoluteUrl(link.Attributes("src").Value, mainUrl)
                    If linkreceived = False Then GoTo jmp

                    'Console.WriteLine("Image: {0}", linkAddress)
                    n = n + 1
chk:

                    If IO.File.Exists("c:\vb\vupics\" + pth + "\" + Trim(Str(n)) + ".jpg") Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        n = n + 1
                        GoTo chk
                    End If
                    Application.DoEvents()
                    'Label1.Text = "Downloading Images " + linkAddress.ToString
                    'Label1.Show()
                    My.Computer.Network.DownloadFile(linkAddress, "c:\vb\vupics\" + pth + "\" + Trim(Str(n) + ".jpg"))
                    If IO.File.Exists("c:\vb\vupics\" + pth + "\" + Trim(Str(n)) + ".jpg") Then
                        ctadd = ctadd + 1
                        Dim pic As New PictureBox
                        pic.ImageLocation = "c:\vb\vupics\" + pth + "\" + Trim(Str(n)) + ".jpg"
                        Dim rnd As New Random
                        pic.Location = New Point(x1, y1)
                        pic.Size = New Size(pic.Width * 2, pic.Height * 2)
                        pic.BackColor = Color.Transparent
                        pic.Visible = True
                        pic.BringToFront()
                        screencords(ctadd) = ctadd
                        sx1(ctadd) = x1
                        sx2(ctadd) = x1 + pic.Width * 2
                        sy1(ctadd) = y1
                        sy2(ctadd) = y1 + pic.Height * 2


                        pic.SizeMode = PictureBoxSizeMode.Zoom
                        hyplink(ctadd) = linkAddress
                        Me.Controls.Add(pic)
                        x1 = x1 + pic.Width
                        If x1 > Me.Width - pic.Width Then
                            x1 = 1
                            y1 = y1 + pic.Height
                        End If





                        Application.DoEvents()
                    End If
jmp:


                Next
            Catch ex As Exception
                GoTo scanimages
            End Try


jmp2:



            'MessageBox.Show(String.Format("Album Art not found for Artist: {0} / Album: {1}", artistName, albumName))

thisjmp:


        Catch ex As Exception
            'MsgBox(ex.ToString)


        End Try
exxxxt:

    End Sub
    Function GetAbsoluteUrl(partialUrl As String, baseUrl As String)
        Try

            Dim myUri = New Uri(partialUrl, UriKind.RelativeOrAbsolute)
            If (myUri.IsAbsoluteUri = False) Then
                myUri = New Uri(New Uri(baseUrl), partialUrl)
            End If
            GetAbsoluteUrl = myUri
            linkreceived = True
        Catch ex As Exception
            linkreceived = False
        End Try

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        x1 = 10 : y1 = 10 : cnt1 = 10 : cnt2 = 10
        TextBox1.Text = "michael jackson"
        TextBox2.Text = "beat it"
    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        msx = e.X
        msy = e.Y

    End Sub
End Class
