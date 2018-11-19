Imports System.Data.OleDb
Public Class MainMDIForm

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximumSize = New Size(My.Computer.Screen.WorkingArea.Size.Width,
                                      My.Computer.Screen.WorkingArea.Size.Height)
        Me.WindowState = FormWindowState.Maximized
        PictureBox4.Width = Me.Width
        PictureBox4.Top = 0
        PictureBox4.Left = 0
        Label1.Width = Me.Width
        MainMenuStrip.Width = Me.Width
        MainMenuStrip.Top = Label1.Height + 1
        PictureBox1.Left = 0
        PictureBox2.Left = Me.Width - 50
        PictureBox2.Top = 10
        PictureBox3.Left = Me.Width - 100
        PictureBox3.Top = 10
        '  PictureBox4.Dock = DockStyle.Top
        MainMenuStrip.Left = 0
        If muser.Equals("super") Then
            RentToolStripMenuItem.Enabled = False
            GodownCloseToolStripMenuItem.Enabled = False
            BillToolStripMenuItem.Enabled = False
            ToolStripMenuItem5.Enabled = False
            ReupdateToolStripMenuItem.Enabled = False
        End If
        Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
        Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
        Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
        Dim XComp As New ADODB.Recordset
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
        Dim rownum As Integer = 0
        Dim MyConn As OleDbConnection
        Dim transaction As OleDbTransaction
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim dag As OleDbDataAdapter
        Dim dsg As DataSet
        Dim dagp As OleDbDataAdapter
        Dim dsgp As DataSet
        Try

            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If

            transaction = MyConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Dim objcmd As New OleDb.OleDbCommand
            Dim objcmdd As New OleDb.OleDbCommand
            objcmd.Connection = MyConn
            objcmd.Transaction = transaction
            objcmd.CommandType = CommandType.Text
            Dim iDate As String
            Dim fDate As DateTime


            iDate = "01/09/2017"
            fDate = Convert.ToDateTime(iDate)
            Dim save As String = "UPDATE [BILL] SET REC_DATE= Format('" & fDate & "','DD/MM/YYYY') WHERE [GROUP]='NEW' AND [GODWN_NO]='056' AND REC_NO='13'"  ''' sorry about that
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()

            iDate = "04/12/2017"
            fDate = Convert.ToDateTime(iDate)
            Dim save1 As String = "UPDATE [BILL] SET REC_DATE= Format('" & fDate & "','DD/MM/YYYY') WHERE [GROUP]='CHALI' AND [GODWN_NO]='042' AND REC_NO='47'"  ''' sorry about that
            objcmd.CommandText = save1
            objcmd.ExecuteNonQuery()
            transaction.Commit()
            MyConn.Close()
        Catch ex As Exception
            MsgBox("Exception: Data update in RECEIPT table in database" & ex.Message)
            Try
                transaction.Rollback()

            Catch
                ' Do nothing here; transaction is not active.
            End Try
        End Try
    End Sub

    Private Sub GroupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GroupToolStripMenuItem.Click


        If Application.OpenForms().OfType(Of FrmGodownType).Any Then
            FrmGodownType.BringToFront()
        Else
            FrmGodownType.Show()
        End If
    End Sub

    Private Sub TenantToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TenantToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmTenant).Any Then
            FrmTenant.BringToFront()
        Else
            FrmTenant.Show()
        End If
    End Sub

    Private Sub BillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BillToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmInvoice).Any Then
            FrmInvoice.BringToFront()
        Else
            FrmInvoice.Show()
        End If
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        Me.Close()
        FrmLogin.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub GodownToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodown).Any Then
            FrmGodown.BringToFront()
        Else
            FrmGodown.Show()
        End If
    End Sub

    Private Sub InvoicePrintingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InvoicePrintingToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If Application.OpenForms().OfType(Of FrmInvSingle).Any Then
            FrmInvSingle.BringToFront()
        Else
            FrmInvSingle.Show()
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RentToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodownTransfer).Any Then
            FrmGodownTransfer.BringToFront()
        Else
            FrmGodownTransfer.Show()
        End If
    End Sub

    Private Sub Format1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Format1ToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmInvSummary).Any Then
            FrmInvSummary.BringToFront()
        Else
            FrmInvSummary.Show()
        End If
    End Sub

    Private Sub Format2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Format2ToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of Form11).Any Then
            Form11.BringToFront()
        Else
            Form11.Show()
        End If
    End Sub

    Private Sub ReceiptToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReceiptToolStripMenuItem.Click

        If Application.OpenForms().OfType(Of FormReceipt).Any Then
            FormReceipt.BringToFront()
        Else
            FormReceipt.Show()
        End If
    End Sub

    Private Sub ToolStripMenuItem3_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If Application.OpenForms().OfType(Of FrmInvoiceSummary).Any Then
            FrmInvoiceSummary.BringToFront()
        Else
            FrmInvoiceSummary.Show()
        End If
    End Sub

    Private Sub GodownCloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownCloseToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodownClose).Any Then
            FrmGodownClose.BringToFront()
        Else
            FrmGodownClose.Show()
        End If
    End Sub

    Private Sub ReceiptPrintingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReceiptPrintingToolStripMenuItem.Click

    End Sub

    Private Sub GSTSpreadsheetsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GSTSpreadsheetsToolStripMenuItem.Click
        'If Application.OpenForms().OfType(Of Form16).Any Then
        '    Form16.BringToFront()
        'Else
        '    Form16.Show()
        'End If
    End Sub

    Private Sub OutstandingReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OutstandingReportToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmRecChecklist).Any Then
            FrmRecChecklist.BringToFront()
        Else
            FrmRecChecklist.Show()
        End If
    End Sub

    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TenantMasterListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TenantMasterListToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmRecChecklist).Any Then
            FrmTenantList.BringToFront()
        Else
            FrmTenantList.Show()
        End If
    End Sub

    Private Sub MonthlyInvoicePrintingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MonthlyInvoicePrintingToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmInvoicePrn).Any Then
            FrmInvoicePrn.BringToFront()
        Else
            FrmInvoicePrn.Show()
        End If
    End Sub

    Private Sub GodownwiseInvoicePrintingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownwiseInvoicePrintingToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmInvoicePrnGdnwise).Any Then
            FrmInvoicePrnGodown.BringToFront()
        Else
            FrmInvoicePrnGodown.Show()
        End If
    End Sub

    Private Sub MonthlyReceiptPrintingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MonthlyReceiptPrintingToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmRecPrn).Any Then
            FrmRecPrn.BringToFront()
        Else
            FrmRecPrn.Show()
        End If
    End Sub

    Private Sub ReceiptPrintingForGodownToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReceiptPrintingForGodownToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmRecPrnGdnwise).Any Then
            FrmRecPrnGdnwise.BringToFront()
        Else
            FrmRecPrnGdnwise.Show()
        End If
    End Sub

    Private Sub TenantsDetailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TenantsDetailToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmPartyDetail).Any Then
            FrmPartyDetail.BringToFront()
        Else
            FrmPartyDetail.Show()
        End If
    End Sub

    Private Sub GodownMasterListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownMasterListToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodownList).Any Then
            FrmGodownList.BringToFront()
        Else
            FrmGodownList.Show()
        End If
    End Sub

    Private Sub TransfterOfGodownListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransfterOfGodownListToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGdTrans).Any Then
            FrmGdTrans.BringToFront()
        Else
            FrmGdTrans.Show()
        End If
    End Sub

    Private Sub GodownCloseListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownCloseListToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGdClose).Any Then
            FrmGdClose.BringToFront()
        Else
            FrmGdClose.Show()
        End If
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        If Application.OpenForms().OfType(Of FrmInvMultiple).Any Then
            FrmInvMultiple.BringToFront()
        Else
            FrmInvMultiple.Show()
        End If
    End Sub

    Private Sub OutstandingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OutstandingToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmOutstanding).Any Then
            FrmOutstanding.BringToFront()
        Else
            FrmOutstanding.Show()
        End If
    End Sub

    Private Sub SummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmSpReport).Any Then
            FrmSpReport.BringToFront()
        Else
            FrmSpReport.Show()
        End If
    End Sub


    Private Sub OpeningAdvanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpeningAdvanceToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmAdvance).Any Then
            FrmAdvance.BringToFront()
        Else
            FrmAdvance.Show()
        End If
    End Sub

    Private Sub ReupdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReupdateToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmReupdate).Any Then
            FrmReupdate.BringToFront()
        Else
            FrmReupdate.Show()
        End If
    End Sub

    Private Sub GodownReopenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownReopenToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodownReopen).Any Then
            FrmGodownReopen.BringToFront()
        Else
            FrmGodownReopen.Show()
        End If
    End Sub

    Private Sub RegenerateBillsWithLogoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegenerateBillsWithLogoToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmInvoiceRegenerate).Any Then
            FrmInvoiceRegenerate.BringToFront()
        Else
            FrmInvoiceRegenerate.Show()
        End If
    End Sub

    Private Sub GSTSpreadsheetsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GSTSpreadsheetsToolStripMenuItem1.Click
        If Application.OpenForms().OfType(Of Form16).Any Then
            Form16.BringToFront()
        Else
            Form16.Show()
        End If
    End Sub

    Private Sub GSTSalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GSTSalesToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGstSales).Any Then
            FrmGstSales.BringToFront()
        Else
            FrmGstSales.Show()
        End If
    End Sub

    Private Sub GSTAdvancesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GSTAdvancesToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGSTAdvance).Any Then
            FrmGSTAdvance.BringToFront()
        Else
            FrmGSTAdvance.Show()
        End If
    End Sub

    Private Sub GodownDetailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GodownDetailToolStripMenuItem.Click
        If Application.OpenForms().OfType(Of FrmGodwnDtl).Any Then
            FrmGodwnDtl.BringToFront()
        Else
            FrmGodwnDtl.Show()
        End If
    End Sub
End Class
