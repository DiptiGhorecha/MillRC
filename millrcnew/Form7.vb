''' <summary>
''' report view form - invoice printing
''' In this module we are displaying .dat file in reachtextbox
''' Initially 1st invoice will disply. Using previous and next button user can view other invoices.
''' </summary>
Public Class Form7
    Private checkPrint As Integer
    Dim formloaded As Boolean = False
    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        checkPrint = RichTextBox1.Print(checkPrint, RichTextBox1.TextLength, e)
        If checkPrint < RichTextBox1.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()    ''''''''close report view form
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '''''set position,size of form and richtextbox
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.Width = Parent.Width - 8
        Me.Height = Parent.Height - 95
        RichTextBox1.Width = Me.Width - 40
        RichTextBox1.Height = Me.Height - 100
        formloaded = True
        '''''If last invoice is visible hide next button
        If FrmInvoicePrn.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrn.TextBox4.Text) Then
            Button3.Enabled = False
        End If
        '''''If 1st invoice is visible hide previous button
        If FrmInvoicePrn.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrn.TextBox3.Text) Then
            Button1.Enabled = False
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ''''''''set current row to next row of datagrid and display next invoice.dat file in richtextbox
        Dim intRow As Integer = FrmInvoicePrn.DataGridView2.CurrentRow.Index
        If intRow < Convert.ToInt32(FrmInvoicePrn.TextBox4.Text) Then
            FrmInvoicePrn.DataGridView2.CurrentRow.Selected = False
            FrmInvoicePrn.DataGridView2.Rows(intRow + 1).Selected = True
            FrmInvoicePrn.DataGridView2.CurrentCell = FrmInvoicePrn.DataGridView2.Rows(intRow + 1).Cells(0)
            Dim flname As String = FrmInvoicePrn.DataGridView2.Item(0, FrmInvoicePrn.DataGridView2.CurrentRow.Index).Value.Replace("/", "_").Replace(" ", "_")
            RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & FrmInvoicePrn.ComboBox4.Text & "\" & FrmInvoicePrn.ComboBox3.Text & "\" & flname & ".dat", RichTextBoxStreamType.PlainText)
            If FrmInvoicePrn.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrn.TextBox4.Text) Then
                Button3.Enabled = False
            End If
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''''''''set current row to previous row of datagrid and display previous invoice.dat file in richtextbox
        Dim intRow As Integer = FrmInvoicePrn.DataGridView2.CurrentRow.Index
        If intRow > Convert.ToInt32(FrmInvoicePrn.TextBox3.Text) Then
            FrmInvoicePrn.DataGridView2.CurrentRow.Selected = False
            FrmInvoicePrn.DataGridView2.Rows(intRow - 1).Selected = True
            FrmInvoicePrn.DataGridView2.CurrentCell = FrmInvoicePrn.DataGridView2.Rows(intRow - 1).Cells(0)
            Dim flname As String = FrmInvoicePrn.DataGridView2.Item(0, FrmInvoicePrn.DataGridView2.CurrentRow.Index).Value.Replace("/", "_").Replace(" ", "_")
            RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & FrmInvoicePrn.ComboBox4.Text & "\" & FrmInvoicePrn.ComboBox3.Text & "\" & flname & ".dat", RichTextBoxStreamType.PlainText)
            If FrmInvoicePrn.DataGridView2.CurrentRow.Index = Convert.ToInt32(FrmInvoicePrn.TextBox3.Text) Then
                Button1.Enabled = False
            End If
            Button3.Enabled = True
        End If
    End Sub

    Private Sub Form7_Move(sender As Object, e As EventArgs) Handles Me.Move
        ''''keep position of form fix
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
End Class