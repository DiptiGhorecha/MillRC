Public Class Form19
    Private checkPrint As Integer
    Dim formloaded As Boolean = False

    Private Sub Form19_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.Width = Parent.Width - 8
        Me.Height = Parent.Height - 95
        RichTextBox1.Width = Me.Width - 40
        RichTextBox1.Height = Me.Height - 100
        formloaded = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim intRow As Integer = FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index
        If intRow < Convert.ToInt32(FrmInvoicePrnGdnwise.TextBox4.Text) Then
            FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Selected = False
            FrmInvoicePrnGdnwise.DataGridView2.Rows(intRow + 1).Selected = True
            FrmInvoicePrnGdnwise.DataGridView2.CurrentCell = FrmInvoicePrnGdnwise.DataGridView2.Rows(intRow + 1).Cells(2)
            Dim flname As String = FrmInvoicePrnGdnwise.DataGridView2.Item(0, FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index).Value.Replace("/", "_").Replace(" ", "_")
            ' Dim flname As String = "0" & FrmInvoicePrnGdnwise.GetValue(FrmInvoicePrnGdnwise.DataGridView2.Item(12, Convert.ToInt32(FrmInvoicePrnGdnwise.TextBox3.Text)).Value).Replace("/", "_").Substring(12, 3)   ' 'TextBox1.Text.Substring(0, 4) & "_" & Month(CDate("1 " & TextBox1.Text.Substring(8, 3))) & "_" & TextBox1.Text.Substring(12, TextBox1.Text.Length - 12)
            RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\" & flname & ".dat", RichTextBoxStreamType.PlainText)
            If FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrnGdnwise.TextBox4.Text) Then
                Button3.Enabled = False
            End If
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim intRow As Integer = FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index
        If intRow > Convert.ToInt32(FrmInvoicePrnGdnwise.TextBox3.Text) Then
            FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Selected = False
            FrmInvoicePrnGdnwise.DataGridView2.Rows(intRow - 1).Selected = True
            FrmInvoicePrnGdnwise.DataGridView2.CurrentCell = FrmInvoicePrnGdnwise.DataGridView2.Rows(intRow - 1).Cells(2)
            Dim flname As String = FrmInvoicePrnGdnwise.DataGridView2.Item(0, FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index).Value.Replace("/", "_").Replace(" ", "_")
            'Dim flname As String = FrmInvoicePrnGdnwise.DataGridView2.Item(0, FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index).Value.Replace("/", "_").Replace(" ", "_")
            RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\" & flname & ".dat", RichTextBoxStreamType.PlainText)
            If FrmInvoicePrnGdnwise.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrnGdnwise.TextBox3.Text) Then
                Button1.Enabled = False
            End If
            Button3.Enabled = True
        End If
    End Sub

    Private Sub Form19_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
End Class