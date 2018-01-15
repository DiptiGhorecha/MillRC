Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class Form16
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
    Public dBaseConnection As New System.Data.OleDb.OleDbConnection
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dagp As OleDbDataAdapter
    Dim dsgp As DataSet
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim fnumm As Integer
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim formloaded As Boolean = False
    Private Sub Form16_Load(sender As Object, e As EventArgs) Handles Me.Load
        'ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        'ComboBox4.Text = DateTime.Now.Year
        'ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        'ComboBox5.Text = DateTime.Now.Year
        If DateTime.Now.Month = 1 Then
            ComboBox3.Text = DateAndTime.MonthName(12)
            ComboBox4.Text = DateTime.Now.Year - 1
            ComboBox2.Text = DateAndTime.MonthName(12)
            ComboBox5.Text = DateTime.Now.Year - 1
        Else
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox4.Text = DateTime.Now.Year
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox5.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox5.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1

        End If
    End Sub
    Private Sub ShowData(mnth As String, yr As String, mnth1 As String, yr1 As String)
        '  konek() 'open our connection
        Try
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim CurrD As DateTime = startP

            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            Dim str As String = "SELECT [BILL].*,[PARTY].P_NAME,[PARTY].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].bill_date between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO"

            da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME,[PARTY].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].bill_date between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            'MsgBox(DataGridView2.ColumnCount)
            DataGridView2.Columns(0).Visible = True
            DataGridView2.Columns(2).Visible = True
            DataGridView2.Columns(4).Visible = True
            DataGridView2.Columns(16).Visible = True
            DataGridView2.Columns(1).Visible = True
            DataGridView2.Columns(0).HeaderText = "Invoice No."
            DataGridView2.Columns(1).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 120
            DataGridView2.Columns(1).Width = 51
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(16).Width = 300
            DataGridView2.Columns(2).HeaderText = "Godown"
            DataGridView2.Columns(16).HeaderText = "Tenant"
            DataGridView2.Columns(4).HeaderText = "Bill Date"
            DataGridView2.Columns(3).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(7).Visible = False
            DataGridView2.Columns(8).Visible = False
            DataGridView2.Columns(9).Visible = False
            DataGridView2.Columns(10).Visible = False
            DataGridView2.Columns(11).Visible = False
            DataGridView2.Columns(12).Visible = False
            DataGridView2.Columns(14).Visible = False
            DataGridView2.Columns(13).Visible = False
            DataGridView2.Columns(15).Visible = False
            DataGridView2.Columns(4).Width = 80
            For X As Integer = 0 To DataGridView2.ColumnCount - 1
                DataGridView2.Columns(X).ReadOnly = True
            Next
            'DataGridView2.Rows(1).Selected = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        End If
        If Trim(ComboBox5.Text) = "" Then
            ComboBox5.Text = DateTime.Now.Year
        End If

        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        If DataGridView1.RowCount > 1 Then
            TextBox1.Text = DataGridView1.Item(4, 0).Value
            TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox5.Text) = "" Then
            ComboBox5.Text = DateTime.Now.Year
        End If
        '        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        If DataGridView1.RowCount > 1 Then
            TextBox1.Text = DataGridView1.Item(4, 0).Value
            TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value

        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox5.Text) = "" Then
            ComboBox5.Text = DateTime.Now.Year
        End If
        '        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        If DataGridView1.RowCount > 1 Then
            TextBox1.Text = DataGridView1.Item(4, 0).Value
            TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value

        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox5.Text) = "" Then
            ComboBox5.Text = DateTime.Now.Year
        End If

        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        If DataGridView1.RowCount > 1 Then
            TextBox1.Text = DataGridView1.Item(4, 0).Value
            TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for selected month")
            ComboBox3.Focus()
            Exit Sub
        End If
        Dim objPRNSetup = New clsPrinterSetup
        'set Paper Lines and Left Margin
        prnmaxpagelines = objPRNSetup.LinesPerPage
        If objPRNSetup.PageSize = PRNA4Paper Then
            prnleftmargin = 7
        Else
            prnleftmargin = 2
        End If
        Dim xcountadj As Integer
        Dim xcountwadj As Integer
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No.'''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        xcountadj = 0
        xcountwadj = 0
        Dim srno As Integer = 0
        'xpage = 1
        xpage = Val("2")
        Dim i1 As Double
        ''''''''''''''''' open a file sharereg.txt'''''''''''
        ' FileOpen(fnum, Application.StartupPath & "\Invoices\RecordSlipView.dat", OpenMode.Output)
        '  Call header()
        Dim numRec As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\gstb2b.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\gstb2b.csv", OpenMode.Output)
        'Dim strbill As Int32 = Convert.ToInt32(TextBox1.Text)
        'Dim edbill As Int32 = Convert.ToInt32(TextBox2.Text)
        Dim strbill As Int32 = 0
        Dim edbill As Int32 = DataGridView2.RowCount - 1
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        If ComboBox1.Text = "All" Then
            chkrs1.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = False
        Else
            chkrs1.Open("SELECT * FROM GST where GST_DESC='" & ComboBox1.Text & "' ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = True
        End If
        HSNRadio1.Checked = True
        Dim HSN As String
        Dim HSNPRINT As Boolean
        Dim totnet, tottaxable, groupnet, grouptaxable As Double
        Dim adtotnet, adtottaxable, adgroupnet, adgrouptaxable As Double
        Dim partyGST As String = ""
        Do While chkrs1.EOF = False
            HSN = chkrs1.Fields(0).Value
            HSNPRINT = True
            For X As Integer = strbill To edbill
                If xcount = 0 Then

                    Print(fnum, GetStringToPrint(16, "GSTIN/UIN of", "S") & GetStringToPrint(35, "Invoice Number", "S") & GetStringToPrint(13, "Invoice Date", "S") & GetStringToPrint(10, " Invoice", "S") & GetStringToPrint(12, " Place of", "S") & GetStringToPrint(7, "Reverse", "S") & GetStringToPrint(12, " Invoice", "S") & GetStringToPrint(10, "E-Commerce GSTIN", "S") & GetStringToPrint(7, "Rate", "N") & GetStringToPrint(15, " Taxable Value", "S") & GetStringToPrint(15, "Cess Amount", "S") & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(16, "recipient", "S") & GetStringToPrint(35, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, " Value", "S") & GetStringToPrint(12, " Supply", "S") & GetStringToPrint(7, "Charge", "S") & GetStringToPrint(12, " Type ", "S") & GetStringToPrint(10, "GSTIN", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, " ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(16, "GSTIN/UIN of", "S") & "," & GetStringToPrint(35, "Invoice Number", "S") & "," & GetStringToPrint(13, "Invoice Date", "S") & "," & GetStringToPrint(10, " Invoice", "S") & "," & GetStringToPrint(12, " Place of", "S") & "," & GetStringToPrint(7, "Reverse", "S") & "," & GetStringToPrint(12, " Invoice", "S") & "," & GetStringToPrint(10, "E-Commerce GSTIN", "S") & "," & GetStringToPrint(7, "Rate", "N") & "," & GetStringToPrint(15, " Taxable Value", "S") & "," & GetStringToPrint(15, "Cess Amount", "S") & "," & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(16, "recipient", "S") & "," & GetStringToPrint(35, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, " Value", "S") & "," & GetStringToPrint(12, " Supply", "S") & "," & GetStringToPrint(7, "Charge", "S") & "," & GetStringToPrint(12, " Type ", "S") & "," & GetStringToPrint(10, "GSTIN", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                    Print(fnum, " " & vbNewLine)
                    Print(fnumm, " " & vbNewLine)

                    xcount = xcount + 1
                End If
                If HSNRadio1.Checked = True Then
                    If DataGridView2.Item(11, X).Value = HSN Then
                        partyGST = ""
                        If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                            partyGST = ""
                        Else
                            partyGST = DataGridView2.Item(17, X).Value
                        End If
                        If B2BRadio1.Checked = True Then
                            If partyGST.Trim <> "" Then
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                    xcount = xcount + 1
                                End If
                                If DataGridView2.Item(15, X).Value = True Then    'b2b
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    adtotnet = adtotnet + DataGridView2.Item(10, X).Value
                                    adgroupnet = adgroupnet + DataGridView2.Item(10, X).Value
                                    adtottaxable = adtottaxable + DataGridView2.Item(5, X).Value
                                    adgrouptaxable = adgrouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                    xcountadj = xcountadj + 1
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                    xcountwadj = xcountwadj + 1
                                End If
                            End If
                        Else
                            If B2BRadio2.Checked = True Then
                                If partyGST.Trim.Equals("") Then
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then   ''''b2c
                                    Else
                                        Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value

                                        xcount = xcount + 1
                                End If
                            End If

                            Else
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                            ' xcount = xcount + 1
                        End If

                    End If
                Else

                    partyGST = ""
                    If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                        partyGST = ""
                    Else
                        partyGST = DataGridView2.Item(17, X).Value
                    End If
                    If B2BRadio1.Checked = True Then
                        If partyGST.Trim <> "" Then
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                            totnet = totnet + DataGridView2.Item(10, X).Value
                            groupnet = groupnet + DataGridView2.Item(10, X).Value
                            tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                            grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                            xcount = xcount + 1
                        End If
                        End If
                    Else
                        If B2BRadio2.Checked = True Then
                            If partyGST.Trim.Equals("") Then
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                xcount = xcount + 1
                            End If
                        End If

                    End If


                End If
            Next
            If HSNRadio1.Checked = True Then
                If chkrs1.EOF = False Then
                    'If groupnet > 0 Then
                    '    Print(fnum, " " & vbNewLine)
                    '    Print(fnumm, " " & vbNewLine)
                    '    Print(fnum, GetStringToPrint(16, "Group Total --> ", "S") & GetStringToPrint(35, xcount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
                    '    Print(fnumm, GetStringToPrint(16, "Group Total --> ", "S") & "," & GetStringToPrint(35, xcount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                    '    groupnet = 0
                    '    grouptaxable = 0
                    'End If
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        chkrs1.Close()
        xcon.Close()
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, xcountwadj, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(totnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, xcountwadj, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(totnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Advance Adj --> ", "S") & GetStringToPrint(35, xcountadj, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(adtotnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(adtottaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Advance Adj --> ", "S") & "," & GetStringToPrint(35, xcountadj, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(adtotnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(adtottaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill start
        Try

            'Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            'Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            'Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            'System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        Dim taxrate As Double = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            Print(fnum, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & GetStringToPrint(35, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(4, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(12, X).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & "," & GetStringToPrint(35, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(4, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(12, X).Value, "S") & vbNewLine)
            groupnet = groupnet + DataGridView1.Item(7, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(4, X).Value
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, StrDup(185, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, StrDup(185, "-") & vbNewLine)

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill end

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill start
        Try

            'Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            'Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            'Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            'System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILL.BILL_DATE;")     '(ACCMST.GST_NO<>'' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            Print(fnum, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & "," & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            groupnet = groupnet + DataGridView1.Item(27, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(19, X).Value
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, StrDup(185, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, StrDup(185, "-") & vbNewLine)

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill end



        ''''''tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO IS NULL order by BILL.BILL_DATE;")     '(ACCMST.GST_NO='' or ACCMST.GST_NO=' ' OR ACCMST.GST_NO='Null') ;")

        FileClose(fnum)
        FileClose(fnumm)
        CreatePDF(Application.StartupPath & "\Reports\gstb2b.dat", Application.StartupPath & "\Reports\gstb2b")
        '''''''''''''''''''''''''''''''''''''''''''''b2c'''''''''''''''''''''''''''''''''''
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No.'''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xcountadj = 0
        xcountwadj = 0
        Dim firsthead As Integer = 0
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        srno = 0
        'xpage = 1
        xpage = Val("2")

        ''''''''''''''''' open a file sharereg.txt'''''''''''
        numRec = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\gstb2c.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\gstb2c.csv", OpenMode.Output)
        strbill = 0  'Convert.ToInt32(TextBox3.Text)
        edbill = DataGridView2.RowCount - 1   'Convert.ToInt32(TextBox4.Text)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        B2BRadio1.Checked = False
        B2BRadio2.Checked = True
        '  ComboBox1.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
        If ComboBox1.Text = "All" Then
            chkrs1.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = False
        Else
            chkrs1.Open("SELECT * FROM GST where GST_DESC='" & ComboBox1.Text & "' ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = True
        End If
        Dim wtaxable As Double = 0
        Dim vtaxable As Double = 0
        Dim VCOUNT As Integer = 0
        Dim WCOUNT As Integer = 0
        HSNRadio1.Checked = True
        totnet = 0
        tottaxable = 0
        groupnet = 0
        grouptaxable = 0
        adtotnet = 0
        adtottaxable = 0
        adgroupnet = 0
        adgrouptaxable = 0
        partyGST = ""
        Do While chkrs1.EOF = False
            HSN = chkrs1.Fields(0).Value
            HSNPRINT = True
            For X As Integer = strbill To edbill
                If HSNRadio1.Checked = True Then
                    If DataGridView2.Item(11, X).Value = HSN Then
                        partyGST = ""
                        If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                            partyGST = ""
                        Else
                            partyGST = DataGridView2.Item(17, X).Value
                        End If
                        If B2BRadio1.Checked = True Then
                            If partyGST.Trim <> "" Then
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                    xcount = xcount + 1
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If B2BRadio2.Checked = True Then
                                If partyGST.Trim.Equals("") Then
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        '  Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        '  Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        ' xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then
                                        If firsthead = 0 Then

                                            Print(fnum, GetStringToPrint(16, "GSTIN/UIN of", "S") & GetStringToPrint(35, "Invoice Number", "S") & GetStringToPrint(13, "Invoice Date", "S") & GetStringToPrint(10, " Invoice", "S") & GetStringToPrint(12, " Place of", "S") & GetStringToPrint(7, "Reverse", "S") & GetStringToPrint(12, " Invoice", "S") & GetStringToPrint(10, "E-Commerce GSTIN", "S") & GetStringToPrint(7, "Rate", "N") & GetStringToPrint(15, " Taxable Value", "S") & GetStringToPrint(15, "Cess Amount", "S") & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                                            Print(fnum, GetStringToPrint(16, "recipient", "S") & GetStringToPrint(35, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, " Value", "S") & GetStringToPrint(12, " Supply", "S") & GetStringToPrint(7, "Charge", "S") & GetStringToPrint(12, " Type ", "S") & GetStringToPrint(10, "GSTIN", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, " ", "S") & vbNewLine)
                                            Print(fnumm, GetStringToPrint(16, "GSTIN/UIN of", "S") & "," & GetStringToPrint(35, "Invoice Number", "S") & "," & GetStringToPrint(13, "Invoice Date", "S") & "," & GetStringToPrint(10, " Invoice", "S") & "," & GetStringToPrint(12, " Place of", "S") & "," & GetStringToPrint(7, "Reverse", "S") & "," & GetStringToPrint(12, " Invoice", "S") & "," & GetStringToPrint(10, "E-Commerce GSTIN", "S") & "," & GetStringToPrint(7, "Rate", "N") & "," & GetStringToPrint(15, " Taxable Value", "S") & "," & GetStringToPrint(15, "Cess Amount", "S") & "," & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                                            Print(fnumm, GetStringToPrint(16, "recipient", "S") & "," & GetStringToPrint(35, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, " Value", "S") & "," & GetStringToPrint(12, " Supply", "S") & "," & GetStringToPrint(7, "Charge", "S") & "," & GetStringToPrint(12, " Type ", "S") & "," & GetStringToPrint(10, "GSTIN", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                                            Print(fnum, " " & vbNewLine)
                                            Print(fnumm, " " & vbNewLine)

                                            firsthead = firsthead + 1
                                        End If
                                        Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        adtotnet = adtotnet + DataGridView2.Item(10, X).Value
                                        adgroupnet = adgroupnet + DataGridView2.Item(10, X).Value
                                        adtottaxable = adtottaxable + DataGridView2.Item(5, X).Value
                                        adgrouptaxable = adgrouptaxable + DataGridView2.Item(5, X).Value
                                        xcountadj = xcountadj + 1
                                        xcount = xcount + 1
                                    Else

                                        ' If DataGridView2.Item(11, X).Value.Equals("997212") Then


                                        '   Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "test" & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)
                                      '  Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & "test" & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)

                                            totnet = totnet + DataGridView2.Item(10, X).Value
                                            groupnet = groupnet + DataGridView2.Item(10, X).Value
                                            tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                            grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                            xcount = xcount + 1
                                            xcountwadj = xcountwadj + 1
                                        '   End If

                                    End If
                                    End If
                            Else
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                            ' xcount = xcount + 1
                        End If

                    End If
                Else
                    partyGST = ""
                    If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                        partyGST = ""
                    Else
                        partyGST = DataGridView2.Item(17, X).Value
                    End If
                    If B2BRadio1.Checked = True Then
                        If partyGST.Trim <> "" Then
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                xcount = xcount + 1
                            End If
                        End If
                    Else
                        If B2BRadio2.Checked = True Then
                            If partyGST.Trim.Equals("") Then
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    '     Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)
                                    '    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                End If
                            End If
                        Else
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                ''    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)
                                '   Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value

                                xcount = xcount + 1
                        End If
                    End If

                End If
                End If
            Next
            If HSNRadio1.Checked = True Then
                If chkrs1.EOF = False Then
                    If groupnet > 0 Then
                        If HSN.Equals("997211") Then
                            wtaxable = wtaxable + grouptaxable + adgrouptaxable
                            WCOUNT = xcount
                        End If
                        If HSN.Equals("997212") Then
                            vtaxable = vtaxable + grouptaxable
                            VCOUNT = xcount
                        End If
                        groupnet = 0
                        grouptaxable = 0
                        xcount = 0
                    End If
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        chkrs1.Close()
        xcon.Close()
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        '  Print(fnum, GetStringToPrint(6, "Note: ", "S") & GetStringToPrint(60, "Advances are not included in invoice value for type rent", "S") & vbNewLine)
        ' Print(fnumm, GetStringToPrint(6, "Note: ", "S") & GetStringToPrint(60, "Advances are not included in invoice value for type rent", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Advance Adj --> ", "S") & GetStringToPrint(35, xcountadj, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(adtotnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(adtottaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Advance Adj --> ", "S") & "," & GetStringToPrint(35, xcountadj, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(adtotnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(adtottaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(6, "Note: ", "S") & GetStringToPrint(60, "Advances are not included in invoice value for type rent", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(6, "Note: ", "S") & GetStringToPrint(60, "Advances are not included in invoice value for type rent", "S") & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, StrDup(185, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(6, "Type  ", "S") & GetStringToPrint(17, "Place of Supply", "S") & GetStringToPrint(8, "Rate    ", "S") & GetStringToPrint(13, "Taxable Value", "S") & GetStringToPrint(15, "Cess Amount", "S") & GetStringToPrint(27, "E-Commerce GSTIN", "S") & GetStringToPrint(20, "Description", "S") & GetStringToPrint(15, "No. of Bills", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(6, "OE", "S") & GetStringToPrint(17, "24-Gujarat", "S") & GetStringToPrint(8, "18.00", "S") & GetStringToPrint(13, Format(vtaxable, "#########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(27, " ", "S") & GetStringToPrint(20, "RENT", "S") & GetStringToPrint(15, VCOUNT, "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(6, "OE", "S") & "," & GetStringToPrint(35, "24-Gujarat", "S") & "," & GetStringToPrint(8, "18.00", "S") & "," & GetStringToPrint(13, Format(vtaxable, "#########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(27, " ", "S") & "," & GetStringToPrint(20, "RENT", "S") & "," & GetStringToPrint(15, VCOUNT, "S") & vbNewLine)
        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill start
        Try
            'Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            'Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            'Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            'System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim CurrD As DateTime = startP
            Dim tb As DataTable
            tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO IS NULL order by BILLWR.BILL_DATE;")     '(ACCMST.GST_NO='' or ACCMST.GST_NO=' ' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            groupnet = groupnet + DataGridView1.Item(7, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(4, X).Value
        Next
        Print(fnum, GetStringToPrint(6, "OE", "S") & GetStringToPrint(17, "24-Gujarat", "S") & GetStringToPrint(8, "18.00", "S") & GetStringToPrint(13, Format(grouptaxable, "#########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(27, " ", "S") & GetStringToPrint(20, "WAREHOUSE", "S") & GetStringToPrint(15, DataGridView1.RowCount, "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(6, "OE", "S") & "," & GetStringToPrint(35, "24-Gujarat", "S") & "," & GetStringToPrint(8, "18.00", "S") & "," & GetStringToPrint(13, Format(grouptaxable, "#########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(27, " ", "S") & "," & GetStringToPrint(20, "WAREHOUSE", "S") & "," & GetStringToPrint(15, DataGridView1.RowCount, "S") & vbNewLine)
        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill end

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill gandalal start
        Try
            'Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            'Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            'Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            'System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim CurrD As DateTime = startP
            Dim tb As DataTable
            tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO IS NULL order by BILL.BILL_DATE;")     '(ACCMST.GST_NO='' or ACCMST.GST_NO=' ' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            groupnet = groupnet + DataGridView1.Item(27, X).Value
            If DataGridView1.Item(19, X).Value < 400 Then
                grouptaxable = grouptaxable + 400
            Else
                grouptaxable = grouptaxable + DataGridView1.Item(19, X).Value
            End If

        Next
        Print(fnum, GetStringToPrint(6, "OE", "S") & GetStringToPrint(17, "24-Gujarat", "S") & GetStringToPrint(8, "18.00", "S") & GetStringToPrint(13, Format(grouptaxable, "#########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(27, " ", "S") & GetStringToPrint(20, "WAREHOUSE GANDALAL", "S") & GetStringToPrint(15, DataGridView1.RowCount, "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(6, "OE", "S") & "," & GetStringToPrint(35, "24-Gujarat", "S") & "," & GetStringToPrint(8, "18.00", "S") & "," & GetStringToPrint(13, Format(grouptaxable, "#########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(27, " ", "S") & "," & GetStringToPrint(20, "WAREHOUSE GANDALAL", "S") & "," & GetStringToPrint(15, DataGridView1.RowCount, "S") & vbNewLine)
        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill gandalal end
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(6, " ", "S") & GetStringToPrint(17, "24-Gujarat", "S") & GetStringToPrint(8, "0.00", "S") & GetStringToPrint(13, Format(wtaxable, "#########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(27, " ", "S") & GetStringToPrint(20, "RESIDENTIAL (NO TAX)", "S") & GetStringToPrint(15, WCOUNT, "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(6, " ", "S") & "," & GetStringToPrint(35, "24-Gujarat", "S") & "," & GetStringToPrint(8, "0.00", "S") & "," & GetStringToPrint(13, Format(wtaxable, "#########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(27, " ", "S") & "," & GetStringToPrint(20, "RESIDENTIAL (NO TAX)", "S") & "," & GetStringToPrint(15, WCOUNT, "S") & vbNewLine)
        FileClose(fnum)
        FileClose(fnumm)
        CreatePDF(Application.StartupPath & "\Reports\gstb2c.dat", Application.StartupPath & "\Reports\gstb2c")

        ''''''''''''''''''''''''''''''''''''''''''''''3rd Spreadsheet''''''''''''''''''''''''''''''''''''start
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No.'''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        srno = 0
        'xpage = 1
        xpage = Val("2")
        numRec = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\gstb3.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\gstb3.csv", OpenMode.Output)
        Dim startnum As String = DataGridView2.Item(0, 0).Value
        Dim endnum As String = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
        Dim count As Integer = DataGridView2.RowCount - 1

        Try
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim tb As DataTable
            tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE=#" & expenddt & "# order by BILLWR.BILL_DATE+BILLWR.BILL_NO;")     '(ACCMST.GST_NO='' or ACCMST.GST_NO=' ' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim wstartnum As String = ""
        Dim wendnum As String = ""
        Dim wcount1 As Integer = 0
        If DataGridView1.RowCount > 0 Then
            wstartnum = DataGridView1.Item(2, 0).Value
            wendnum = DataGridView1.Item(2, DataGridView1.RowCount - 1).Value
            wcount1 = DataGridView1.RowCount
        End If
        Try
            Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text.ToString, ComboBox3.SelectedIndex + 1)
            Dim edate As String = DaysInMonth.ToString & "/" & String.Format("{0:00}", ComboBox3.SelectedIndex + 1) & "/" & ComboBox4.Text.ToString
            Dim expenddt As Date = Date.ParseExact(edate, "dd/MM/yyyy",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Dim tb As DataTable
            tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE=#" & expenddt & "#;")     '(ACCMST.GST_NO='' or ACCMST.GST_NO=' ' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim vstartnum As String = ""
        Dim vendnum As String = ""
        Dim vcount1 As Integer = 0

        If DataGridView1.RowCount > 0 Then
            vstartnum = DataGridView1.Item(2, 0).Value
            vendnum = DataGridView1.Item(2, DataGridView1.RowCount - 1).Value
            vcount1 = DataGridView1.RowCount
        End If
        Print(fnum, GetStringToPrint(57, "Summary of documents issued during the tax period (13) ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Summary of documents issued during the tax period (13) ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, "Total Number ", "S") & GetStringToPrint(20, "Total Cancelled ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, "Total Number ", "S") & "," & GetStringToPrint(20, "Total Cancelled ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, count + vcount1 + wcount1, "N") & GetStringToPrint(20, "0", "N") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, count + vcount1 + wcount1, "N") & "," & GetStringToPrint(20, "0", "N") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Nature of Document", "S") & GetStringToPrint(20, "Sr. No. From", "S") & GetStringToPrint(20, "Sr. No. To ", "S") & GetStringToPrint(20, "Total Number", "S") & GetStringToPrint(20, "Cancelled", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Nature of Document", "S") & "," & GetStringToPrint(20, "Sr. No. From ", "S") & "," & GetStringToPrint(20, "Sr. No. To ", "S") & "," & GetStringToPrint(20, "Total Number", "S") & "," & GetStringToPrint(20, "Cancelled", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Invoice for outward supply", "S") & GetStringToPrint(20, IIf(count > 0, "GO-" & startnum, ""), "S") & GetStringToPrint(20, IIf(count > 0, "GO-" & endnum, ""), "S") & GetStringToPrint(20, count, "S") & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Invoice for outward supply", "S") & "," & GetStringToPrint(20, IIf(count > 0, "GO-" & startnum, ""), "S") & "," & GetStringToPrint(20, IIf(count > 0, "GO-" & endnum, ""), "S") & "," & GetStringToPrint(20, count, "S") & "," & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Invoice for outward supply", "S") & GetStringToPrint(20, IIf(wcount1 > 0, "WF-" & wstartnum, ""), "S") & GetStringToPrint(20, IIf(wcount1 > 0, "WF-" & wendnum, ""), "S") & GetStringToPrint(20, wcount1, "S") & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Invoice for outward supply", "S") & "," & GetStringToPrint(20, IIf(wcount1 > 0, "WF-" & wstartnum, ""), "S") & "," & GetStringToPrint(20, IIf(wcount1 > 0, "WF-" & wendnum, ""), "S") & "," & GetStringToPrint(20, wcount1, "S") & "," & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Invoice for outward supply", "S") & GetStringToPrint(20, IIf(vcount1 > 0, "WV-" & vstartnum, ""), "S") & GetStringToPrint(20, IIf(vcount1 > 0, "WV-" & vendnum, ""), "S") & GetStringToPrint(20, vcount1, "S") & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Invoice for outward supply", "S") & "," & GetStringToPrint(20, IIf(vcount1 > 0, "WV-" & vstartnum, ""), "S") & "," & GetStringToPrint(20, IIf(vcount1 > 0, "WV-" & vendnum, ""), "S") & "," & GetStringToPrint(20, vcount1, "S") & "," & GetStringToPrint(20, "0", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Debit Note", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Debit Note", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Debit Note", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Debit Note", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Delivery Challan for job work", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Delivery Challan for job work", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Invoice for inward supply from unregistered person", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Invoice for inward supply from unregistered person", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Refund Voucher", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Refund Voucher", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(57, "Invoice for outward supply", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & GetStringToPrint(20, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(57, "Invoice for outward supply", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & "," & GetStringToPrint(20, " ", "S") & vbNewLine)
        FileClose(fnum)
        FileClose(fnumm)
        CreatePDF(Application.StartupPath & "\Reports\gstb3.dat", Application.StartupPath & "\Reports\gstb3")
        ''''''''''''''''''''''''''''''''''''''''''''''3rd spreadsheet end''''''''''''''''''''''''''''''''end

        ''''''''''''''''''''''''''''''''''''''''''''b2c''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''advance adjusted'''''''''''''''''''''''''''''''''''''''
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No.'''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        srno = 0
        'xpage = 1
        xpage = Val("2")
        ''''''''''''''''' open a file sharereg.txt'''''''''''
        ' FileOpen(fnum, Application.StartupPath & "\Invoices\RecordSlipView.dat", OpenMode.Output)
        '  Call header()
        numRec = 0
        totnet = 0
        tottaxable = 0
        groupnet = 0
        grouptaxable = 0
        adtotnet = 0
        adtottaxable = 0
        adgroupnet = 0
        adgrouptaxable = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\atadj.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\atadj.csv", OpenMode.Output)
        'Dim strbill As Int32 = Convert.ToInt32(TextBox1.Text)
        'Dim edbill As Int32 = Convert.ToInt32(TextBox2.Text)
        strbill = 0
        edbill = DataGridView2.RowCount - 1
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        If ComboBox1.Text = "All" Then
            chkrs1.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = False
        Else
            chkrs1.Open("SELECT * FROM GST where GST_DESC='" & ComboBox1.Text & "' ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = True
        End If
        HSNRadio1.Checked = True
        B2BRadio3.Checked = True
        B2BRadio1.Checked = False
        B2BRadio2.Checked = False
        partyGST = ""
        Do While chkrs1.EOF = False
            HSN = chkrs1.Fields(0).Value
            HSNPRINT = True
            For X As Integer = strbill To edbill
                If xcount = 0 Then

                    Print(fnum, GetStringToPrint(16, "GSTIN/UIN of", "S") & GetStringToPrint(35, "Invoice Number", "S") & GetStringToPrint(13, "Invoice Date", "S") & GetStringToPrint(10, " Invoice", "S") & GetStringToPrint(12, " Place of", "S") & GetStringToPrint(7, "Reverse", "S") & GetStringToPrint(12, " Invoice", "S") & GetStringToPrint(10, "E-Commerce GSTIN", "S") & GetStringToPrint(7, "Rate", "N") & GetStringToPrint(15, " Taxable Value", "S") & GetStringToPrint(15, "Cess Amount", "S") & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(16, "recipient", "S") & GetStringToPrint(35, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, " Value", "S") & GetStringToPrint(12, " Supply", "S") & GetStringToPrint(7, "Charge", "S") & GetStringToPrint(12, " Type ", "S") & GetStringToPrint(10, "GSTIN", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, " ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(16, "GSTIN/UIN of", "S") & "," & GetStringToPrint(35, "Invoice Number", "S") & "," & GetStringToPrint(13, "Invoice Date", "S") & "," & GetStringToPrint(10, " Invoice", "S") & "," & GetStringToPrint(12, " Place of", "S") & "," & GetStringToPrint(7, "Reverse", "S") & "," & GetStringToPrint(12, " Invoice", "S") & "," & GetStringToPrint(10, "E-Commerce GSTIN", "S") & "," & GetStringToPrint(7, "Rate", "N") & "," & GetStringToPrint(15, " Taxable Value", "S") & "," & GetStringToPrint(15, "Cess Amount", "S") & "," & GetStringToPrint(55, "Tenant Name", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(16, "recipient", "S") & "," & GetStringToPrint(35, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, " Value", "S") & "," & GetStringToPrint(12, " Supply", "S") & "," & GetStringToPrint(7, "Charge", "S") & "," & GetStringToPrint(12, " Type ", "S") & "," & GetStringToPrint(10, "GSTIN", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                    Print(fnum, " " & vbNewLine)
                    Print(fnumm, " " & vbNewLine)

                    xcount = xcount + 1
                End If
                If HSNRadio1.Checked = True Then
                    If DataGridView2.Item(11, X).Value = HSN Then
                        partyGST = ""
                        If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                            partyGST = ""
                        Else
                            partyGST = DataGridView2.Item(17, X).Value
                        End If
                        If B2BRadio1.Checked = True Then
                            If partyGST.Trim <> "" Then
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    '   Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    '   Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                    xcount = xcount + 1
                                End If
                                If DataGridView2.Item(15, X).Value = True Then    'b2b
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If B2BRadio2.Checked = True Then
                                If partyGST.Trim.Equals("") Then
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then   ''''b2c
                                    Else
                                        Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value

                                        xcount = xcount + 1
                                    End If
                                End If

                            Else
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    '  Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    '  Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                    If DataGridView2.Item(11, X).Value.Equals("997212") Then

                                        Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value

                                        xcount = xcount + 1
                                    End If
                                Else
                                        'Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        'Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                        'totnet = totnet + DataGridView2.Item(10, X).Value
                                        'groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        'tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        'grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                        'xcount = xcount + 1
                                    End If
                            End If
                            ' xcount = xcount + 1
                        End If

                    End If
                Else

                    partyGST = ""
                    If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                        partyGST = ""
                    Else
                        partyGST = DataGridView2.Item(17, X).Value
                    End If
                    If B2BRadio1.Checked = True Then
                        If partyGST.Trim <> "" Then
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                xcount = xcount + 1
                            End If
                        End If
                    Else
                        If B2BRadio2.Checked = True Then
                            If partyGST.Trim.Equals("") Then
                                If DataGridView2.Item(15, X).Value = True Then
                                Else
                                    Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If DataGridView2.Item(15, X).Value = True Then
                            Else
                                Print(fnum, GetStringToPrint(16, partyGST, "S") & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(16, partyGST, "S") & "," & GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                xcount = xcount + 1
                            End If
                        End If

                    End If


                End If
            Next
            If HSNRadio1.Checked = True Then
                If chkrs1.EOF = False Then
                    'If groupnet > 0 Then
                    '    Print(fnum, " " & vbNewLine)
                    '    Print(fnumm, " " & vbNewLine)
                    '    Print(fnum, GetStringToPrint(16, "Group Total --> ", "S") & GetStringToPrint(35, xcount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
                    '    Print(fnumm, GetStringToPrint(16, "Group Total --> ", "S") & "," & GetStringToPrint(35, xcount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                    '    groupnet = 0
                    '    grouptaxable = 0
                    'End If
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        chkrs1.Close()
        xcon.Close()
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, xcount - 1, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(totnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, xcount - 1, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(totnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(36, "Place of Supply", "S") & GetStringToPrint(10, "Rate", "S") & GetStringToPrint(33, "Gross Advance Adjusted", "S") & GetStringToPrint(20, " Cess Amount", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(36, "Place of Supply", "S") & "," & GetStringToPrint(10, "Rate", "S") & "," & GetStringToPrint(33, "Gross Advance Adjusted", "S") & "," & GetStringToPrint(20, " Cess Amount", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(36, "24-Gujarat", "S") & GetStringToPrint(36, "18", "S"))
        Print(fnumm, GetStringToPrint(36, "24-Gujarat", "S") & "," & GetStringToPrint(36, "18", "S") & ",")
        Print(fnum, GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & GetStringToPrint(14, Format(tottaxable * 18 / 100, "##########0.00"), "N"))
        Print(fnumm, GetStringToPrint(14, Format(tottaxable, "##########0.00"), "N") & "," & GetStringToPrint(14, Format(tottaxable * 18 / 100, "##########0.00"), "N"))
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

        FileClose(fnum)
        FileClose(fnumm)
        CreatePDF(Application.StartupPath & "\Reports\atadj.dat", Application.StartupPath & "\Reports\atadj")


        ''''''''''''''''''''''''''advance adjusted'''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''advance receipt'''''''''''''''''''''''''''''''''''''
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        numRec = 0
        Dim xline As Integer = 0

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim DysInMonth As Integer = 1
        '  MsgBox(DaysInMonth)
        Dim strtP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DysInMonth)
        DysInMonth = Date.DaysInMonth(ComboBox5.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim edP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DysInMonth)
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        numRec = 0
        xline = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\at.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\at.csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & strtP & "','DD/MM/YYYY') AND FORMAT('" & edP & "','DD/MM/YYYY')  order by YEAR([RECEIPT].REC_DATE)+[RECEIPT].REC_NO"

        chkrs1.Open(str, xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim totamt As Double = 0
        Dim tot1, tot2, tot3, tot4 As Double
        tot1 = 0
        tot2 = 0
        tot3 = 0
        tot4 = 0
        Dim totadv As Double = 0
        Do While chkrs1.EOF = False

            If first Then
                Print(fnum, GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(15, "  Advance Rent", "S") & GetStringToPrint(15, "Advance SGST", "S") & GetStringToPrint(15, "Advance CGST", "S") & GetStringToPrint(15, "Total Advance", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(15, "  Advance Rent", "S") & "," & GetStringToPrint(15, "Advance SGST", "S") & "," & GetStringToPrint(15, "Advance CGST", "S") & "," & GetStringToPrint(15, "Total Advance", "S") & vbNewLine)
                Print(fnum, StrDup(110, "=") & vbNewLine)
                Print(fnumm, StrDup(110, "=") & vbNewLine)
                first = False
                    xline = xline + 3
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
            Dim residential As Boolean = False
            If chkrs4.Fields(37).Value.Equals("997211") Then
                residential = True
            End If

            Dim census As String = ""
            Dim survey As String = ""
            Dim pname As String = ""
            Dim rent As Double = 0
            Dim pcode1 As String = ""
            Dim hsnm As String = ""
            Dim CGST_RATE As Double = 0
            Dim SGST_RATE As Double = 0
            Dim CGST_TAXAMT As Double = 0
            Dim SGST_TAXAMT As Double = 0
            Dim gst As Double = 0
            Dim gst_amt As Double = 0
            Dim net As Double
            Dim rnd As Integer
            If chkrs4.EOF = False Then
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                    Else
                        census = chkrs4.Fields(5).Value
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                    Else
                        survey = chkrs4.Fields(4).Value
                    End If
                    pname = chkrs4.Fields(38).Value
                    pcode1 = chkrs4.Fields(1).Value

                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    Dim amtt As Double = 0
                    If chkrs2.EOF = False Then
                        chkrs2.MoveFirst()
                        amtt = chkrs2.Fields(4).Value
                        If IsDBNull(chkrs2.Fields(5).Value) Then
                        Else
                            amtt = amtt + chkrs2.Fields(5).Value
                        End If
                    End If
                    chkrs2.Close()
                    chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs4.Fields(37).Value & "'", xcon)

                    If chkrs3.EOF = False Then
                        If IsDBNull(chkrs3.Fields(2).Value) Then
                            CGST_RATE = 0
                        Else
                            CGST_RATE = chkrs3.Fields(2).Value
                        End If
                        If IsDBNull(chkrs3.Fields(3).Value) Then
                            SGST_RATE = 0
                        Else
                            SGST_RATE = chkrs3.Fields(3).Value
                        End If
                    End If
                    gst = CGST_RATE + SGST_RATE
                    chkrs3.Close()
                    gst_amt = gst * amtt / 100
                    rnd = gst_amt - Math.Round(gst_amt)
                    If rnd >= 50 Then
                        gst_amt = Math.Round(gst_amt) + 1
                    Else
                        gst_amt = Math.Round(gst_amt)
                    End If

                    net = amtt + gst_amt
                    CGST_TAXAMT = gst_amt / 2
                'CGST_TAXAMT = amt * CGST_RATE / 100
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    'SGST_TAXAMT = amt * SGST_RATE / 100
                    SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                End If
                chkrs4.Close()

                Dim grp As String = chkrs1.Fields(1).Value
                Dim gdn As String = chkrs1.Fields(2).Value
                Dim invdt As DateTime = chkrs1.Fields(3).Value
                Dim inv As Integer = chkrs1.Fields(4).Value
                Dim FIRSTREC As Boolean = True
                Dim FROMNO As String = ""
                Dim TONO As String
                Dim against As String = ""
                Dim against1 As String = ""
                Dim against3 As String = ""
                Dim against2 As String = ""
            Dim agcount As Integer = 0
            Dim adjusted_amt As Double = 0

            '  If chkrs1.Fields(6).Value = True Then

            '  Else
            '  grp = chkrs1.Fields(1).Value
            '   gdn = chkrs1.Fields(2).Value
            Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False
                    'sgsrate = chkrs2.Fields(8).Value
                    'cgsrate = chkrs2.Fields(6).Value
                    'sgamt = chkrs2.Fields(9).Value
                    'cgamt = chkrs2.Fields(7).Value
                    If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                        If FIRSTREC Then
                            FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                            TONO = FROMNO
                            FIRSTREC = False
                        Else
                            TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        End If
                        pname = chkrs2.Fields(16).Value
                        adjusted_amt = adjusted_amt + chkrs2.Fields(10).Value
                        If agcount < 8 Then
                            against = against + "GO-" & chkrs2.Fields(0).Value & ", "
                        Else
                            If agcount < 15 Then
                                against1 = against1 + "GO-" & chkrs2.Fields(0).Value & ", "
                            Else
                                If agcount < 22 Then
                                    against2 = against2 + "GO-" & chkrs2.Fields(0).Value & ", "
                                Else
                                    If agcount < 29 Then
                                        against3 = against3 + "GO-" & chkrs2.Fields(0).Value & ", "
                                    Else
                                    End If
                                End If
                            End If
                        End If

                        agcount = agcount + 1
                    End If
                    If chkrs2.EOF = False Then
                        chkrs2.MoveNext()
                    End If
                    If chkrs2.EOF = True Then
                        Exit Do
                    End If

                Loop
                chkrs2.Close()
                '   End If
                If against.Length > 2 Then
                    against = against.Substring(0, against.Length - 2)
                End If
                If Trim(against1).Equals("") Then
                Else
                    against1 = against1.Substring(0, against1.Length - 2)
                End If
                If Trim(against2).Equals("") Then
                Else
                    against2 = against2.Substring(0, against2.Length - 2)
                End If
                If Trim(against3).Equals("") Then
                Else
                    against3 = against3.Substring(0, against3.Length - 2)
                End If
                '''''''''''''''''''against bill and period end''''''''''''
                '    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)

                ''''''''''''''find out if any advance is left after adjustment start
                Dim advanceamt As Double = 0
                Dim advanceamtprint As Double = 0
                Dim lastbilladjusted As Integer
                Dim last_bldate As DateTime
                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                advanceamtprint = advanceamt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs5.Open(Rss, xcon)
                    Do While chkrs5.EOF = False
                        '  chkrs5.MoveLast()
                        If chkrs1.Fields(3).Value >= chkrs5.Fields(4).Value Then
                            lastbilladjusted = chkrs5.Fields(0).Value
                            last_bldate = chkrs5.Fields(4).Value
                        End If
                        If chkrs5.EOF = False Then
                            chkrs5.MoveNext()
                        End If
                    Loop
                    chkrs5.Close()
                    If lastbilladjusted = 0 Then
                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND [STATUS]='C' AND P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                        chkrs5.Open(Rss, xcon)
                        If chkrs5.EOF = False Then
                            ' chkrs5.MoveLast()
                            'lastbilladjusted = chkrs5.Fields(13).Value
                            last_bldate = chkrs5.Fields(0).Value
                        End If
                        chkrs5.Close()
                    End If
                    Dim dtcounter As Integer = 1
                    Do Until advanceamt <= 0
                        If FIRSTREC Then
                            If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                FROMNO = MonthName(Convert.ToDateTime(last_bldate).Month, False) & "-" & Convert.ToDateTime(last_bldate).Year
                                advanceamt = advanceamt - net
                                TONO = FROMNO
                                FIRSTREC = False
                                'last_bldate = chkrs5.Fields(0).Value
                            End If
                        Else
                            TONO = MonthName(Convert.ToDateTime(last_bldate).AddMonths(dtcounter).Month, False) & "-" & Convert.ToDateTime(last_bldate).AddMonths(dtcounter).Year
                            advanceamt = advanceamt - net
                            dtcounter = dtcounter + 1
                        End If
                    Loop
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                totamt = totamt + chkrs1.Fields(5).Value
                totadv = totadv + advanceamtprint
                Dim adv As Boolean = True
                If adv = True Then
                If chkrs1.Fields(6).Value = True And residential = False Then
                    Print(fnum, GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(15, Format((advanceamtprint * 100 / 118), "######0.00"), "N") & GetStringToPrint(15, Format((advanceamtprint * 100 / 118) * 9 / 100, "######0.00"), "N") & GetStringToPrint(15, Format((advanceamtprint * 100 / 118) * 9 / 100, "######0.00"), "N") & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & vbNewLine)
                    Print(fnumm, GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(15, Format((advanceamtprint * 100 / 118), "######0.00"), "N") & "," & GetStringToPrint(15, Format((advanceamtprint * 100 / 118) * 9 / 100, "######0.00"), "N") & "," & GetStringToPrint(15, Format((advanceamtprint * 100 / 118) * 9 / 100, "######0.00"), "N") & "," & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & vbNewLine)
                    tot1 = tot1 + (advanceamtprint * 100 / 118)
                    tot2 = tot2 + ((advanceamtprint * 100 / 118) * 9 / 100)
                    tot3 = tot2 + ((advanceamtprint * 100 / 118) * 9 / 100)
                    tot4 = tot4 + advanceamtprint
                    xline = xline + 1
                End If
            Else
                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)

                    xline = xline + 1
                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    xline = xline + 1

                    Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                    xline = xline + 1
                    If Trim(against1).Equals("") Then
                    Else
                        '    against1 = against1.Substring(0, against1.Length - 2)
                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                        xline = xline + 1
                    End If
                    If Trim(against2).Equals("") Then
                    Else
                        '   against2 = against2.Substring(0, against2.Length - 2)
                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                        xline = xline + 1
                    End If
                    If Trim(against3).Equals("") Then
                    Else
                        '  against3 = against3.Substring(0, against3.Length - 2)
                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                        xline = xline + 1
                    End If
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    xline = xline + 1
                End If
                '    Print(fnum, StrDup(80, "-") & vbNewLine)
                '  Next
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If

        Loop
        Print(fnum, StrDup(180, "-") & vbNewLine)
        '  Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
        Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(13, Format((tot1), "######0.00"), "N") & GetStringToPrint(15, Format(tot2, "######0.00"), "N") & GetStringToPrint(15, Format(tot3, "######0.00"), "N") & GetStringToPrint(15, Format(tot4, "######0.00"), "N") & vbNewLine)
        Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(13, Format((tot1), "######0.00"), "N") & "," & GetStringToPrint(15, Format(tot2, "######0.00"), "N") & "," & GetStringToPrint(15, Format(tot3, "######0.00"), "N") & "," & GetStringToPrint(15, Format(tot4, "######0.00"), "N") & vbNewLine)
        Print(fnum, StrDup(110, "-") & vbNewLine)
        Print(fnumm, StrDup(110, "-") & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(36, "Place of Supply", "S") & GetStringToPrint(10, "Rate", "S") & GetStringToPrint(33, "Gross Advance Received", "S") & GetStringToPrint(20, " Cess Amount", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(36, "Place of Supply", "S") & "," & GetStringToPrint(10, "Rate", "S") & "," & GetStringToPrint(33, "Gross Advance Received", "S") & "," & GetStringToPrint(20, " Cess Amount", "S") & vbNewLine)
        Print(fnum, StrDup(110, "=") & vbNewLine)
        Print(fnumm, StrDup(110, "=") & vbNewLine)
        Print(fnum, GetStringToPrint(36, "24-Gujarat", "S") & GetStringToPrint(36, "18", "S"))
        Print(fnumm, GetStringToPrint(36, "24-Gujarat", "S") & "," & GetStringToPrint(36, "18", "S") & ",")
        Print(fnum, GetStringToPrint(14, Format(tot1, "##########0.00"), "N") & GetStringToPrint(14, Format(tot2 + tot3, "##########0.00"), "N"))
        Print(fnumm, GetStringToPrint(14, Format(tot1, "##########0.00"), "N") & "," & GetStringToPrint(14, Format(tot2 + tot3, "##########0.00"), "N"))
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

        chkrs1.Close()
        xcon.Close()
        FileClose(fnum)
        FileClose(fnumm)
        CreatePDF(Application.StartupPath & "\Reports\at.dat", Application.StartupPath & "\Reports\at")
        ''''''''''''''''''''''''''advance receipt'''''''''''''''''''''''''''''''''''''
        MsgBox(Application.StartupPath + "\Reports\gstb2b.CSV file is generated")
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
    Public Function packtable()
        Dim Conn As String
        Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\WHBILL; Extended Properties=dBase IV;"
        dBaseConnection = New System.Data.OleDb.OleDbConnection(Conn)
        Dim cmd As New OleDbCommand()
        cmd.Connection = dBaseConnection
        cmd.CommandType = System.Data.CommandType.Text
        cmd.CommandText = "PACK " & Application.StartupPath & "\WHBILL\BILLWR.dbf"
        dBaseConnection.Open()
        cmd.ExecuteNonQuery()
        dBaseConnection.Close()
        '  MsgBox("PACK complete")
    End Function
    Public Function getdbasetable(ByVal SqlString As String) As DataTable
        Dim ReturnableTable As New DataTable
        Try
            OpendBConnection()
            Dim SelectCommand As New System.Data.OleDb.OleDbCommand(SqlString, dBaseConnection)
            Dim TableAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
            TableAdapter.SelectCommand = SelectCommand
            TableAdapter.Fill(ReturnableTable)
            Return ReturnableTable
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & SqlString, 16, "Error")
            End
        End Try
        Return ReturnableTable
    End Function

    Public Sub OpendBConnection()
        Try
            Dim ConnectionString As String
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\WHBILL; Extended Properties=dBase IV;"
            dBaseConnection = New System.Data.OleDb.OleDbConnection(ConnectionString)

            If dBaseConnection.State = 0 Then dBaseConnection.Open()
        Catch ex As Exception
            MsgBox(ex.Message, 16, "Error")
        End Try
    End Sub

    Private Sub Form16_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
End Class