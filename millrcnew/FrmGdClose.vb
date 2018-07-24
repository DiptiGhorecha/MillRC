Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Public Class FrmGdClose
    Dim formloaded As Boolean = False
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim indexorder As String = "GODWN_NO"
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim fnum As Integer
    Dim fnumm As Integer
    Private Sub FrmGdClose_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        ShowData()
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
        formloaded = True
    End Sub

    Private Sub FrmGdClose_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE order by [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(4).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(3).Width = 71
            DataGridView2.Columns(2).HeaderText = "Godown"
            DataGridView2.Columns(7).HeaderText = "Tenant"
            DataGridView2.Columns(3).HeaderText = "Date"
            '  DataGridView2.Columns(5).HeaderText = "Reason"
            DataGridView2.Columns(7).Width = 250
            DataGridView2.Columns(5).Width = 100

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        If e.ColumnIndex = 7 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
        Else
            indexorder = "[CLGDWN].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (DateTimePicker2.Value.Date < DateTimePicker1.Value.Date) Then
            MsgBox("To Date must be equal or grater than From Date")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Gdclosedlist.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where [CLGDWN].[TO_D] >=format('" & Convert.ToDateTime(DateTimePicker1.Value) & "','dd/mm/yyyy') AND [CLGDWN].[TO_D]<=format('" & Convert.ToDateTime(DateTimePicker2.Value) & "','dd/mm/yyyy') order by [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1

        Do While chkrs1.EOF = False


            If first Then
                Print(fnum, GetStringToPrint(107, "Closed / Suspended Godown List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Type", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(15, "Closing Date", "S") & " " & GetStringToPrint(50, "Tenant Name", "S") & " " & GetStringToPrint(20, "Closed / Suspended", "S") & vbNewLine)
                Print(fnum, StrDup(120, "=") & vbNewLine)
                Print(fnumm, GetStringToPrint(107, "Closed / Suspended Godown List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnumm, StrDup(14, " ") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "Sr. No.", "S") & "," & GetStringToPrint(10, "Godown Type", "S") & "," & GetStringToPrint(10, "Godown Code", "S") & "," & GetStringToPrint(15, "Closing Date", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(20, "Closed / Suspended", "S") & vbNewLine)
                Print(fnumm, StrDup(120, "=") & vbNewLine)
                first = False
                xline = xline + 4
            End If

            Dim cld As String
            If IsDBNull(chkrs1.Fields(6).Value) Or chkrs1.Fields(6).Value.Equals("") Then
                cld = "Closed"
            Else
                cld = "Suspended"
            End If
            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(10, chkrs1.Fields(2).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(3).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & " " & GetStringToPrint(20, cld, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(7, counter, "S") & "," & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(10, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(15, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(20, cld, "S") & vbNewLine)

            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FileClose(fnumm)
        FrmGdCloseView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Gdclosedlist.dat", RichTextBoxStreamType.PlainText)
        FrmGdCloseView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Gdclosedlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
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
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (DateTimePicker2.Value.Date < DateTimePicker1.Value.Date) Then
            MsgBox("To Date must be equal or grater than From Date")
            DateTimePicker1.Focus()
            Exit Sub
        End If

        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Gdclosedlist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where [CLGDWN].[TO_D] >=format('" & Convert.ToDateTime(DateTimePicker1.Value) & "','dd/mm/yyyy') AND [CLGDWN].[TO_D]<=format('" & Convert.ToDateTime(DateTimePicker2.Value) & "','dd/mm/yyyy') order by [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1

        Do While chkrs1.EOF = False


            If first Then
                Print(fnum, GetStringToPrint(107, "Closed / Suspended Godown List From Date : " & DateTimePicker1.Value.ToShortDateString & " to " & DateTimePicker2.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Type", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(15, "Closing Date", "S") & " " & GetStringToPrint(50, "Tenant Name", "S") & " " & GetStringToPrint(20, "Closed / Suspended", "S") & vbNewLine)
                Print(fnum, StrDup(120, "=") & vbNewLine)
                first = False
                xline = xline + 4
            End If

            Dim cld As String
            If IsDBNull(chkrs1.Fields(6).Value) Or chkrs1.Fields(6).Value.Equals("") Then
                cld = "Closed"
            Else
                cld = "Suspended"
            End If
            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(10, chkrs1.Fields(2).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(3).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(7).Value, "S") & " " & GetStringToPrint(20, cld, "S") & vbNewLine)

            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FrmGdCloseView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Gdclosedlist.dat", RichTextBoxStreamType.PlainText)
        FrmGdCloseView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Gdclosedlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub
End Class