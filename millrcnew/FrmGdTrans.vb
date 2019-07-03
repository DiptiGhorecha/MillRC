Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
''' <summary>
''' tables used - gdtrans,party
''' this is form to accept inputs from user to view/print godown transfer master
''' Only report view is complete. For report print not checked alignments and font size
''' FrmGdTransView.vb is used to hold report view 
''' </summary>
Public Class FrmGdTrans
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Private indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim fnum As Integer
    Dim fnumm As Integer
    Private Sub FrmGdTrans_Load(sender As Object, e As EventArgs) Handles Me.Load
        '''''''set position of form
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True

        DateTimePicker1.Value = DateTime.Now    ''''assign current date to from transfer date
        DateTimePicker2.Value = DateTime.Now    ''''assign current date to To transfer date
        ShowData()
        formloaded = True
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        '''''''''search transfer datagrid for the text user type in search text box
        MyConn = New OleDbConnection(connString)
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY date", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        ''''''set order of transfer datagrid as user click on header of datagrid column
        If e.ColumnIndex = 0 Then
            indexorder = "[GROUP]"
            GroupBox5.Text = "Search by Group Type"
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "GODWN_NO"
            GroupBox5.Text = "Search by Godown"
        End If
        If e.ColumnIndex = 6 Then
            indexorder = "[DATE]"
            GroupBox5.Text = "Search by Date"
        End If
        If e.ColumnIndex = 7 Then
            indexorder = "P2.P_NAME"
            GroupBox5.Text = "Search by Old tenant name"
        End If
        If e.ColumnIndex = 8 Then
            indexorder = "P1.P_NAME"
            GroupBox5.Text = "Search by New tenant name"
        End If

    End Sub

    Private Sub FrmGdTrans_Move(sender As Object, e As EventArgs) Handles Me.Move
        '''''''keep position of form fix
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub ShowData()
        '''''show godown transfer detail to transfer datagrid using gdtrans table
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            da = New OleDb.OleDbDataAdapter("SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) order by DATE", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(2).Visible = False
            DataGridView2.Columns(3).Visible = False
            DataGridView2.Columns(4).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(0).Visible = False
            DataGridView2.Columns(7).Visible = False
            DataGridView2.Columns(8).Visible = False
            DataGridView2.Columns(0).Visible = True
            DataGridView2.Columns(1).Visible = True
            DataGridView2.Columns(6).Visible = True
            DataGridView2.Columns(7).Visible = True
            DataGridView2.Columns(8).Visible = True

            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(1).Width = 71
            DataGridView2.Columns(6).Width = 71
            DataGridView2.Columns(7).Width = 200
            DataGridView2.Columns(8).Width = 200
            DataGridView2.Columns(1).HeaderText = "Godown"
            DataGridView2.Columns(6).HeaderText = "Transfer Date"
            DataGridView2.Columns(7).HeaderText = "Old Tenant"
            DataGridView2.Columns(8).HeaderText = "New Tenant"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '''''report view
        If (DateTimePicker2.Value.Date < DateTimePicker1.Value.Date) Then
            MsgBox("To Transfer Date must be equal or grater than From Transfer Date")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2   ''''''''for .csv file'''''''''''''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Gdtranslist.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If

        Dim str As String = "SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) where [GDTRANS].[DATE]>=format('" & Convert.ToDateTime(DateTimePicker1.Value.ToShortDateString) & "','dd/mm/yyyy') AND [GDTRANS].[DATE]<=format('" & Convert.ToDateTime(DateTimePicker2.Value.ToShortDateString) & "','dd/mm/yyyy')  order by Date"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1

        Do While chkrs1.EOF = False


            If first Then

                Print(fnum, GetStringToPrint(107, "Godown Transfer List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Type", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(15, "Transfer Date", "S") & " " & GetStringToPrint(50, "Old Tenant Name", "S") & " " & GetStringToPrint(50, "New Tenant Name", "S") & vbNewLine)
                Print(fnum, StrDup(130, "=") & vbNewLine)
                Print(fnumm, GetStringToPrint(107, "Godown Transfer List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnumm, StrDup(14, " ") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "Sr. No.", "S") & "," & GetStringToPrint(10, "Godown Type", "S") & "," & GetStringToPrint(10, "Godown Code", "S") & "," & GetStringToPrint(15, "Transfer Date", "S") & "," & GetStringToPrint(50, "Old Tenant Name", "S") & "," & GetStringToPrint(50, "New Tenant Name", "S") & vbNewLine)
                Print(fnumm, StrDup(130, "=") & vbNewLine)
                first = False
                xline = xline + 5
            End If

            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(10, chkrs1.Fields(1).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(6).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(7, counter, "S") & "," & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(10, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(15, chkrs1.Fields(6).Value, "S") & "," & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & vbNewLine)

            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FileClose(fnumm)
        '''''''display created .dat file to richtextbox of report view form
        FrmGdTransView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Gdtranslist.dat", RichTextBoxStreamType.PlainText)
        ''''''show report
        FrmGdTransView.Show()
        ''''''convert pdf file from dat file
        CreatePDF(Application.StartupPath & "\Reports\Gdtranslist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        '''''''convert pdf file from dat file
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 15
            pdfPage.Width = 842
            pdfPage.Height = 595
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 8, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1
                line = readFile.ReadLine()
                If counter >= 43 Then
                    counter = 0
                    pdfPage = pdf.AddPage()
                    graph = XGraphics.FromPdfPage(pdfPage)
                    font = New XFont("COURIER NEW", 8, XFontStyle.Regular)

                    pdfPage.TrimMargins.Left = 15

                    pdfPage.Width = 842
                    pdfPage.Height = 595
                    yPoint = 60
                End If
                If line Is Nothing Then
                    Exit While
                Else
                    graph.DrawString(line, font, XBrushes.Black,
                    New XRect(30, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = invoice_no & ".pdf"


            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''''''''report printing
        If (DateTimePicker2.Value.Date < DateTimePicker1.Value.Date) Then
            MsgBox("To Transfer Date must be equal or grater than From Transfer Date")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Gdtranslist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) where [GDTRANS].[DATE]>=format('" & Convert.ToDateTime(DateTimePicker1.Value.ToShortDateString) & "','dd/mm/yyyy') AND [GDTRANS].[DATE]<=format('" & Convert.ToDateTime(DateTimePicker2.Value.ToShortDateString) & "','dd/mm/yyyy')  order by Date"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1

        Do While chkrs1.EOF = False


            If first Then
                Print(fnum, GetStringToPrint(107, "Godown Transfer List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Type", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(15, "Transfer Date", "S") & " " & GetStringToPrint(50, "Old Tenant Name", "S") & " " & GetStringToPrint(50, "New Tenant Name", "S") & vbNewLine)
                Print(fnum, StrDup(130, "=") & vbNewLine)
                first = False
                xline = xline + 5
            End If


            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(10, chkrs1.Fields(1).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(6).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & vbNewLine)

            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FrmGdTransView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Gdtranslist.dat", RichTextBoxStreamType.PlainText)
        FrmGdTransView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Gdtranslist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)

        ''''send pdf file to default printer
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()   ''''close the form
    End Sub
End Class