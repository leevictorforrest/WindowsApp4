Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices

Public Class Form1
    Public linkreceived As Boolean
    Public x1, x2, y1, y2 As Integer
    Public cnt1, cnt2 As Integer
    Public ctadd As Long
    Public screencordsx(1024) As Integer
    Public screencordsy(1024) As Integer
    Public screencordsx1(1024) As Integer
    Public screencordsy1(1024) As Integer
    Public sx1(1024) As Integer
    Public sx2(1024) As Integer
    Public sy1(1024) As Integer
    Public sy2(1024) As Integer
    Public hyplink(1024) As String
    Public msx, msy As Long
    Public patho As String
    Public ccc As Long
    Public onelongx, onelongy As Long
    Public sumofpics As Long
    Public pic(1000) As PictureBox
    Public hyperlnk(1000) As String
    public hyperlnknum as long

    Public Class testmouseclass

        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetAsyncKeyState(ByVal vkey As Long) As Long

        End Function


        Public Shared Function LeftMouseIsDown() As Boolean
            try
                'Return GetAsyncKeyState(Keys.LButton) > 0 And &H8000
            Catch ex As Exception

            End Try
            if keys.LButton=true then
                LeftMouseIsDown=true
                MsgBox("pressed")
            End If
           return LeftMouseIsDown

        End Function

    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Enabled = False
        GetAlbumInfo()

        ' "C:\vb\vupics\Ftorrents100 Greatest Classic Rock Songs (2019)001.  Van Halen  -  Jump (2015 Remaster).mp3\1.jpg"

    End Sub

    Public Structure POINTAPI
        Dim x As Long
        Dim y As Long
    End Structure


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Dim cur_x As Single
        Dim cur_y As Single

        cur_x = GetMousePoint.x
        cur_y = GetMousePoint.y
        'me.TextBox3.text  = cur_x & "   " & cur_y


        Timer1.Enabled = True
    End Sub
    Public Function GetMousePoint() As POINTAPI




        Dim fx = Me.Width
        Dim fy = Me.Height
        Dim plx As Long
        Dim ply As Long
        Dim sp As Integer
        Dim mmmc As Boolean
        Dim num As Integer
        For g = 1 To ctadd - 1
            Dim x As Long = Cursor.Position.X
            Dim y As Long = Cursor.Position.Y

           
            fx = pic(g).Left
            fy = pic(g).Top
            plx = pic(g).Right
            ply = pic(g).Bottom
            If x >= fx And x <= plx And y >= fy And y <= ply Then
                sp = g
                GoTo jp
            End If
        Next
        GoTo joppa
jp:
        Try
         
            Dim newImage As Image = pic(sp).Image
            hyperlnknum=sp
            Dim strea As New MemoryStream
            newImage.Save(strea, ImageFormat.Png)
            strea.Position = 0
            Dim gi As Image = Image.FromStream(strea)
            Dim g1111 As Graphics
            g1111 = Graphics.FromHwnd(Form2.Handle)
            Dim iw, ih As Long
            iw = newImage.Width
            ih = newImage.Height
            TextBox2.Text = Str(fx) + " " + Str(fy) + " " + Str(sp)
            'dim pb as new PictureBox 
            'pb.image=newimage
            Dim point As Point = pic(sp).PointToClient(Control.MousePosition)
            g1111.SmoothingMode = SmoothingMode.AntiAlias
            g1111.Clear(Color.White)
            For c = 1 To 8
                For a = 1 To 360 Step Math.PI * 96
                    Dim x1 As Double = c * Math.Cos(Math.PI * a / 180.0) + point.X



                    Dim y1 As Double = c * Math.Sin(Math.PI * a / 180.0) + point.Y

                    g1111.DrawImage(gi, CInt(x1), CInt(y1), Form2.Width - iw, Form2.Height - ih)


                Next a

            Next c
            g1111.DrawImage(newImage, CInt(x1), CInt(y1), x1 + Form2.Width - iw, y1 + Form2.Height - ih)
          




        Catch ex As Exception

        End Try


joppa:








    End Function

    Private Function FindPointOnCircle(originPoint As Point, radius As Double, angleDegrees As Double) As Point




        Dim x As Double = radius * Math.Cos(Math.PI * angleDegrees / 180.0) + originPoint.X



        Dim y As Double = radius * Math.Sin(Math.PI * angleDegrees / 180.0) + originPoint.Y








        Return New Point(x, y)




    End Function



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
        For xx = 0 To ctadd
            For Each cControl As Control In Me.Controls



                If cControl.Name = "" Then
                    Me.Controls.Remove(cControl)
                    screencordsx(xx) = 0
                    screencordsx(xx) = 0
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
        Me.WindowState = FormWindowState.Maximized

        Form2.Show

        'MsgBox(zArtist + vbLf + zTrackTitle + vbLf + zAlbum)
        For Each dr In My.Computer.FileSystem.GetDirectories("c:\vb\vupics\", FileIO.SearchOption.SearchAllSubDirectories)
            Dim fnb As Long = 0
            For Each fil In My.Computer.FileSystem.GetFiles(dr, FileIO.SearchOption.SearchAllSubDirectories, "*.jpg")


                ctadd = ctadd + 1
                fnb = fnb + 1
                Dim tf As String = IO.File.ReadAllText(dr + "\" + Trim(Str(fnb)) + ".txt")
                Dim im As New PictureBox

                If ctadd > 120 Then GoTo exxxt



                pic(ctadd) = New PictureBox
                pic(ctadd).SizeMode = PictureBoxSizeMode.Zoom
                pic(ctadd).ImageLocation = tf

                sx1(ctadd) = x1
                sx2(ctadd) = pic(ctadd).Width
                sy1(ctadd) = y1
                sy2(ctadd) = pic(ctadd).Height

                Dim rnd As New Random
                pic(ctadd).Location = New Point(x1, y1)

                im.ImageLocation = tf
                im.Location = New Point(x1, y1)
                Me.Controls.Add(im)

                'pic(ctadd).Size = New Size(sx2(ctadd), sy2(ctadd))
                pic(ctadd).BackColor = Color.Transparent
                pic(ctadd).Visible = True



                sumofpics = sumofpics + x1 + pic(ctadd).Width / y1 + pic(ctadd).Height


                'pic(ctadd).SizeMode=

                hyplink(ctadd) = tf
                hyperlnk(ctadd) = tf


                Me.Controls.Add(pic(ctadd))

                AddHandler pic(ctadd).MouseMove, AddressOf GetMousePoint
                AddHandler pic(ctadd).MouseClick, AddressOf GetMousePoint

                Dim blackPen As New Pen(Color.Black, 2)
                Dim myGraphics As Graphics

                myGraphics = Graphics.FromHwnd(Me.Handle)


                ' Create rectangle. 
                Dim rect As New Rectangle(x1, y1, x1 + pic(ctadd).Width, y1 + pic(ctadd).Height)
                Dim rect2 As New Rectangle(sx1(ctadd), sy1(ctadd), sx2(ctadd), sy2(ctadd))


                ' Draw rectangle to screen.
                myGraphics.DrawRectangle(blackPen, rect2)

                pic(ctadd).BringToFront()

                Cursor.Position = New Point(x1, y1)
                screencordsx(ctadd) = pic(ctadd).Location.X
                screencordsy(ctadd) = pic(ctadd).Location.Y
                screencordsx1(ctadd) = pic(ctadd).Width
                screencordsy1(ctadd) = pic(ctadd).Height

                Dim rect3 As New Rectangle(screencordsx(ctadd), screencordsy(ctadd), screencordsx1(ctadd), screencordsy1(ctadd))

                Dim blackPen2 As New Pen(Color.Black, 8)


                myGraphics.DrawRectangle(blackPen2, rect3)
                x1 = x1 + pic(ctadd).Width

                If x1 > Me.Width - pic(ctadd).Width Then
                    x1 = 1
                    y1 = y1 + pic(ctadd).Height
                End If


                Application.DoEvents()


            Next
        Next
exxxt:



        Timer1.Enabled = True
        Form2.Show
    End Sub

    Private Sub Panel1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
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
        onelongx = sumofpics - (msx / 120)
        onelongy = sumofpics - (msy / 120)

        TextBox3.Text = ""
        TextBox1.Text = Str(onelongx) + " " + Str(onelongy)
        TextBox2.Text = Str(msx) + " " + Str(msy) + " dif " + Str(onelongx - msx) + " " + Str(onelongy - msy)
        Timer1.Enabled = True

    End Sub

    Private Sub Form1_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        Process.Start(me.hyperlnk(me.hyperlnknum))
    End Sub
End Class
