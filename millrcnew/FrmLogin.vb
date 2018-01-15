Public Class FrmLogin
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Equals("super") Or TextBox1.Text.Equals("SUPER") Then
            muser = "super"
            Me.Hide()
            MainMDIForm.Show()
        Else
            If TextBox1.Text.Equals("abc") Or TextBox1.Text.Equals("ABC") Then
                muser = "abc"
                Me.Hide()
                MainMDIForm.Show()
            Else
                MsgBox("Invalid Password")
                Me.Close()
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp
        If e.KeyCode = 13 Then
            Button1_Click(Me, EventArgs.Empty)
        End If
    End Sub
End Class