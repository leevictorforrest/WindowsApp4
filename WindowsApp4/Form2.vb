Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form2_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        Process.Start(form1.hyperlnk(form1.hyperlnknum))
    End Sub
End Class