Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing

Public Class FrmGodownList
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"

    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim dap As OleDbDataAdapter
    Dim dsp As DataSet
    Dim dat As OleDbDataAdapter
    Dim dst As DataSet
    Dim dar As OleDbDataAdapter
    Dim dsr As DataSet
    Dim dagv As OleDbDataAdapter
    Dim dsgv As DataSet
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim strReportFilePath As String
    Dim GrpAddCorrect As String
    Dim blnTranStart As Boolean
    Dim oldName As String
    Dim ok As Boolean
    Private bValidategodown As Boolean = True
    Private bValidatetype As Boolean = True
    Private bValidatepname As Boolean = True
    Private bValidateHSN As Boolean = True
    Private bValidaterent As Boolean = True
    Dim formloaded As Boolean = False
    Private indexorder As String = "[PARTY].P_NAME"
    Private frmload As Boolean = True
    Private tabrec As Integer = 0
    Private colnum As Integer = 0
    Private rownum As Integer = 0
    Dim ctrlname As String = "TextBox1"
    Dim fnum As Integer
    Dim fnumm As Integer
    Private ststatus As String

    Private Sub FrmGodownList_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False

            fillgroupcombo()
            ShowData()

            formloaded = True
            If DataGridView2.RowCount > 1 Then
                TextBox1.Text = DataGridView2.Item(3, 0).Value
                TextBox2.Text = DataGridView2.Item(3, DataGridView2.RowCount - 1).Value
            End If
            ComboBox2.Text = "Current"
            ststatus = "C"
            '  TextBox1.Text = chkrs1.Fields(1).Value
        Catch ex As Exception
            MessageBox.Show("Error loading form godown master list " & ex.Message)
        End Try
    End Sub
    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate
        If ComboBox1.FindString(ComboBox1.Text) < 0 Then
            ComboBox1.Text = ComboBox1.Text.Remove(ComboBox1.Text.Length - 1)
            ComboBox1.SelectionStart = ComboBox1.Text.Length
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        ctrlname = "TextBox1"
    End Sub
    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        ctrlname = "TextBox2"
    End Sub
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            If ComboBox1.Text.Equals("") Then
                da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            Else
                da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [godown].[group]='" & ComboBox1.Text & "' order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            End If

            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection
            'For i As Integer = 0 To DataGridView2.Columns.Count - 1
            '    DataGridView2.Columns(i).Visible = False
            'Next
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(2).Visible = False
            DataGridView2.Columns(4).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(7).Visible = False
            DataGridView2.Columns(8).Visible = False
            DataGridView2.Columns(9).Visible = False
            DataGridView2.Columns(10).Visible = False
            DataGridView2.Columns(11).Visible = False
            DataGridView2.Columns(12).Visible = False
            DataGridView2.Columns(13).Visible = False
            DataGridView2.Columns(14).Visible = False
            DataGridView2.Columns(15).Visible = False
            DataGridView2.Columns(16).Visible = False
            DataGridView2.Columns(17).Visible = False
            DataGridView2.Columns(18).Visible = False
            DataGridView2.Columns(19).Visible = False
            DataGridView2.Columns(20).Visible = False
            DataGridView2.Columns(21).Visible = False
            DataGridView2.Columns(22).Visible = False
            DataGridView2.Columns(23).Visible = False
            DataGridView2.Columns(24).Visible = False
            DataGridView2.Columns(25).Visible = False
            DataGridView2.Columns(26).Visible = False
            DataGridView2.Columns(27).Visible = False
            DataGridView2.Columns(28).Visible = False
            DataGridView2.Columns(29).Visible = False
            DataGridView2.Columns(30).Visible = False
            DataGridView2.Columns(31).Visible = False
            DataGridView2.Columns(32).Visible = False
            DataGridView2.Columns(33).Visible = False
            DataGridView2.Columns(34).Visible = False
            DataGridView2.Columns(35).Visible = False
            DataGridView2.Columns(36).Visible = False
            DataGridView2.Columns(37).Visible = False
            DataGridView2.Columns(0).Visible = True
            'DataGridView2.Columns(0).HeaderCell.
            DataGridView2.Columns(3).Visible = True
            DataGridView2.Columns(38).Visible = True
            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(3).Width = 71
            DataGridView2.Columns(38).Width = 405
            DataGridView2.Columns(3).HeaderText = "Godown"
            DataGridView2.Columns(38).HeaderText = "Tenant"
            DataGridView2.Columns(21).HeaderText = "Outstanding"
            ' DataGridView2.Columns(21).Visible = True
            DataGridView2.Columns(21).Width = 105

            'DataGridView2.Rows(1).Selected = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function fillgroupcombo()
        Try
            '  Dim authors As New AutoCompleteStringCollection
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dag = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] Order by [GROUP].G_CODE", MyConn)
            dsg = New DataSet
            dsg.Clear()
            dag.Fill(dsg, "GROUP")
            ComboBox1.DataSource = dsg.Tables("GROUP")
            ComboBox1.DisplayMember = "G_CODE"
            ComboBox1.ValueMember = "G_CODE"
            dag.Dispose()
            dsg.Dispose()
            MyConn.Close() ' close connection
            For i = 0 To ComboBox1.Items.Count - 1
                '      authors.Add(ComboBox1.Items(i).ToString)
            Next i
            '    ComboBox1.AutoCompleteMode = AutoCompleteMode.Suggest
            'ComboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
            '        ComboBox1.AutoCompleteCustomSource = authors
        Catch ex As Exception
            MessageBox.Show("Group combo fill :" & ex.Message)
        End Try
    End Function
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        Dim i As Integer = DataGridView2.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(3, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView2.Item(3, i).Value)
        End If
        '  GroupBox5.Visible = False
        '  DataGridView2.Visible = False
        ' ' Me.Width = Me.Width - DataGridView2.Width + 15
        '  Me.Height = Me.Height - 145
        If ctrlname = "TextBox1" Then
            TextBox2.Focus()
        Else
            If ctrlname = "TextBox2" Then
                Button1.Focus()
            Else
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ShowData()
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(3, 0).Value
            TextBox2.Text = DataGridView2.Item(3, DataGridView2.RowCount - 1).Value
        End If
    End Sub

    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' and [GODOWN].[GROUP]='" & ComboBox1.Text & "' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")
        DataGridView2.DataSource = ds.Tables("GODOWN")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub

    Private Sub FrmGodownList_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter from godown no.")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter to godown no.")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As String = TextBox1.Text
        Dim edrec As String = TextBox2.Text

        If edrec < strrec Then
            MsgBox("From godown no. must be less than To godown no.")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Godownmasterlist.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [godown].[group]='" & ComboBox1.Text & "' and [GODOWN].[GODWN_NO]>='" & strrec & "' AND [GODOWN].[GODWN_NO]<='" & edrec & "' AND [GODOWN].[STATUS]='" & ststatus & "' order by [GODOWN].GROUP+[GODOWN].GODWN_NO"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1
        Dim ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GST, REMARK As String
        Do While chkrs1.EOF = False


            If first Then
                'globalHeader("Godown Master List", fnum, fnumm)
                Print(fnum, GetStringToPrint(57, ComboBox2.Text & " Godown Master List for type - " & chkrs1.Fields(0).Value, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(50, "Tenant Name", "S") & " " & GetStringToPrint(15, "Using From", "S") & " " & GetStringToPrint(12, "Survey No.", "S") & " " & GetStringToPrint(20, "Godown Size", "S") & " " & GetStringToPrint(10, "HSN", "S") & " " & GetStringToPrint(9, "Rent", "N") & " " & GetStringToPrint(9, "Remarks", "N") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "Census No.", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "Pe Rent", "N") & vbNewLine)
                Print(fnum, StrDup(210, "=") & vbNewLine)

                Print(fnumm, GetStringToPrint(57, ComboBox2.Text & " Godown Master List for type - " & chkrs1.Fields(0).Value, "S") & vbNewLine)
                Print(fnumm, StrDup(14, " ") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "Sr. No.", "S") & "," & GetStringToPrint(10, "Godown Code", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(15, "Using From", "S") & "," & GetStringToPrint(12, "Survey No.", "S") & "," & GetStringToPrint(20, "Godown Size", "S") & "," & GetStringToPrint(10, "HSN", "S") & "," & GetStringToPrint(9, "Rent", "N") & "," & GetStringToPrint(9, "Remarks", "N") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(15, "", "S") & "," & GetStringToPrint(12, "Census No.", "S") & "," & GetStringToPrint(20, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(9, "Pe Rent", "N") & vbNewLine)
                Print(fnumm, StrDup(210, "=") & vbNewLine)
                first = False
                xline = xline + 5
            End If

            Dim gdsize As String = ""
            gdsize = chkrs1.Fields(18).Value.ToString & " * " & chkrs1.Fields(19).Value.ToString & " = " & chkrs1.Fields(20).Value.ToString
            chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(0).Value & "' and GODWN_NO='" & chkrs1.Fields(3).Value & "' and P_CODE ='" & chkrs1.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
            Dim amt As Double = 0
            Dim amt1 As Double = 0
            If chkrs2.EOF = False Then
                chkrs2.MoveFirst()

                amt = chkrs2.Fields(4).Value
                If IsDBNull(chkrs2.Fields(5).Value) Then
                Else
                    amt1 = chkrs2.Fields(5).Value
                End If
            End If
            chkrs2.Close()

            Dim addArr() As String
            Dim rem1 As String = ""
            Dim rem2 As String = ""
            Dim rem3 As String = ""
            Dim rem4 As String = ""

            If IsDBNull(chkrs1.Fields(22).Value) Or chkrs1.Fields(22).Value.Equals("") Then

            Else
                rem1 = chkrs1.Fields(22).Value
            End If
            If IsDBNull(chkrs1.Fields(23).Value) Or chkrs1.Fields(23).Value.Equals("") Then

            Else
                rem2 = chkrs1.Fields(23).Value
            End If
            If IsDBNull(chkrs1.Fields(24).Value) Or chkrs1.Fields(24).Value.Equals("") Then

            Else
                rem3 = chkrs1.Fields(24).Value
            End If
            If IsDBNull(chkrs1.Fields(25).Value) Or chkrs1.Fields(25).Value.Equals("") Then

            Else
                rem4 = chkrs1.Fields(25).Value
            End If
            If (rem1.IndexOf(vbLf) >= 0) Then
                addArr = rem1.Split(vbLf)
                rem1 = addArr(0)
                If addArr.Length > 1 Then
                    rem2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                    rem3 = addArr(2)
                End If
                If addArr.Length > 3 Then
                    rem4 = addArr(3)
                End If
            End If

            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(3).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(38).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(11).Value, "S") & " " & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(4).Value), " ", chkrs1.Fields(4).Value), "S") & " " & GetStringToPrint(20, gdsize, "S") & " " & GetStringToPrint(10, chkrs1.Fields(37).Value, "S") & " " & GetStringToPrint(9, Format(amt, "#####0.00"), "N") & " " & GetStringToPrint(65, rem1, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(7, counter, "S") & "," & GetStringToPrint(10, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(50, chkrs1.Fields(38).Value, "S") & "," & GetStringToPrint(15, chkrs1.Fields(11).Value, "S") & "," & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(4).Value), " ", chkrs1.Fields(4).Value), "S") & "," & GetStringToPrint(20, gdsize, "S") & "," & GetStringToPrint(10, chkrs1.Fields(37).Value, "S") & "," & GetStringToPrint(9, Format(amt, "#####0.00"), "N") & "," & GetStringToPrint(65, rem1, "S") & vbNewLine)
            If IsDBNull(amt1) Or amt1 = 0 Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(5).Value), " ", chkrs1.Fields(5).Value), "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, Format(amt1, "#####0.00"), "N") & " " & GetStringToPrint(65, rem2, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(15, "", "S") & "," & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(5).Value), " ", chkrs1.Fields(5).Value), "S") & "," & GetStringToPrint(20, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(9, Format(amt1, "#####0.00"), "N") & "," & GetStringToPrint(65, rem2, "S") & vbNewLine)
                counter = counter + 1
            End If
            If IsDBNull(rem3) Or rem3.Equals("") Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "", "N") & " " & GetStringToPrint(65, rem3, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(15, "", "S") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(20, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(9, "", "N") & "," & GetStringToPrint(65, rem3, "S") & vbNewLine)
                counter = counter + 1
            End If
            If IsDBNull(rem4) Or rem4.Equals("") Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "", "N") & " " & GetStringToPrint(65, rem4, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(7, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(15, "", "S") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(20, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(9, "", "N") & "," & GetStringToPrint(65, rem4, "S") & vbNewLine)
                counter = counter + 1
            End If
            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FileClose(fnumm)
        FrmGodownListView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Godownmasterlist.dat", RichTextBoxStreamType.PlainText)
        FrmGodownListView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Godownmasterlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)

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
            Dim font As XFont = New XFont("COURIER NEW", 6, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1
                line = readFile.ReadLine()
                If counter >= 43 Then
                    counter = 0
                    pdfPage = pdf.AddPage()
                    graph = XGraphics.FromPdfPage(pdfPage)
                    font = New XFont("COURIER NEW", 6, XFontStyle.Regular)

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter from godown no.")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter to godown no.")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As String = TextBox1.Text
        Dim edrec As String = TextBox2.Text

        If edrec < strrec Then
            MsgBox("From godown no. must be less than To godown no.")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Godownmasterlist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [godown].[group]='" & ComboBox1.Text & "' and [GODOWN].[GODWN_NO]>='" & strrec & "' AND [GODOWN].[GODWN_NO]<='" & edrec & "' AND [GODOWN].[STATUS]='" & ststatus & "' order by [GODOWN].GROUP+[GODOWN].GODWN_NO"
        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1
        Dim ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GST, REMARK As String
        Do While chkrs1.EOF = False


            If first Then
                globalHeader("Godown Master List", fnum, 0)
                Print(fnum, GetStringToPrint(57, ComboBox2.Text & " Godown Master List for type - " & chkrs1.Fields(0).Value, "S") & vbNewLine)
                Print(fnum, StrDup(14, " ") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(10, "Godown Code", "S") & " " & GetStringToPrint(50, "Tenant Name", "S") & " " & GetStringToPrint(15, "Using From", "S") & " " & GetStringToPrint(12, "Survey No.", "S") & " " & GetStringToPrint(20, "Godown Size", "S") & " " & GetStringToPrint(10, "HSN", "S") & " " & GetStringToPrint(9, "Rent", "N") & " " & GetStringToPrint(9, "Remarks", "N") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "Census No.", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "Pe Rent", "N") & vbNewLine)
                Print(fnum, StrDup(210, "=") & vbNewLine)
                first = False
                xline = xline + 5
            End If

            Dim gdsize As String = ""
            gdsize = chkrs1.Fields(18).Value.ToString & " * " & chkrs1.Fields(19).Value.ToString & " = " & chkrs1.Fields(20).Value.ToString
            chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(0).Value & "' and GODWN_NO='" & chkrs1.Fields(3).Value & "' and P_CODE ='" & chkrs1.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
            Dim amt As Double = 0
            Dim amt1 As Double = 0
            If chkrs2.EOF = False Then
                chkrs2.MoveFirst()

                amt = chkrs2.Fields(4).Value
                If IsDBNull(chkrs2.Fields(5).Value) Then
                Else
                    amt1 = chkrs2.Fields(5).Value
                End If
            End If
            chkrs2.Close()

            Dim addArr() As String
            Dim rem1 As String = ""
            Dim rem2 As String = ""
            Dim rem3 As String = ""
            Dim rem4 As String = ""

            If IsDBNull(chkrs1.Fields(22).Value) Or chkrs1.Fields(22).Value.Equals("") Then

            Else
                rem1 = chkrs1.Fields(22).Value
            End If
            If IsDBNull(chkrs1.Fields(23).Value) Or chkrs1.Fields(23).Value.Equals("") Then

            Else
                rem2 = chkrs1.Fields(23).Value
            End If
            If IsDBNull(chkrs1.Fields(24).Value) Or chkrs1.Fields(24).Value.Equals("") Then

            Else
                rem3 = chkrs1.Fields(24).Value
            End If
            If IsDBNull(chkrs1.Fields(25).Value) Or chkrs1.Fields(25).Value.Equals("") Then

            Else
                rem4 = chkrs1.Fields(25).Value
            End If
            If (rem1.IndexOf(vbLf) >= 0) Then
                addArr = rem1.Split(vbLf)
                rem1 = addArr(0)
                If addArr.Length > 1 Then
                    rem2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                    rem3 = addArr(2)
                End If
                If addArr.Length > 3 Then
                    rem4 = addArr(3)
                End If
            End If

            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(10, chkrs1.Fields(3).Value, "S") & " " & GetStringToPrint(50, chkrs1.Fields(38).Value, "S") & " " & GetStringToPrint(15, chkrs1.Fields(11).Value, "S") & " " & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(4).Value), " ", chkrs1.Fields(4).Value), "S") & " " & GetStringToPrint(20, gdsize, "S") & " " & GetStringToPrint(10, chkrs1.Fields(37).Value, "S") & " " & GetStringToPrint(9, Format(amt, "#####0.00"), "N") & " " & GetStringToPrint(65, rem1, "S") & vbNewLine)
            If IsDBNull(amt1) Or amt1 = 0 Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, IIf(IsDBNull(chkrs1.Fields(5).Value), " ", chkrs1.Fields(5).Value), "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, Format(amt1, "#####0.00"), "N") & " " & GetStringToPrint(65, rem2, "S") & vbNewLine)
                counter = counter + 1
            End If
            If IsDBNull(rem3) Or rem3.Equals("") Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "", "N") & " " & GetStringToPrint(65, rem3, "S") & vbNewLine)
                counter = counter + 1
            End If
            If IsDBNull(rem4) Or rem4.Equals("") Then
            Else
                Print(fnum, GetStringToPrint(7, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(50, "", "S") & " " & GetStringToPrint(15, "", "S") & " " & GetStringToPrint(12, "", "S") & " " & GetStringToPrint(20, "", "S") & " " & GetStringToPrint(10, "", "S") & " " & GetStringToPrint(9, "", "N") & " " & GetStringToPrint(65, rem4, "S") & vbNewLine)
                counter = counter + 1
            End If
            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FrmGodownListView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Godownmasterlist.dat", RichTextBoxStreamType.PlainText)
        FrmGodownListView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Godownmasterlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text.Equals("Current") Then
            ststatus = "C"
        Else
            If ComboBox2.Text.Equals("Closed") Then
                ststatus = "D"
            Else
                ststatus = "S"
            End If
        End If
    End Sub
End Class