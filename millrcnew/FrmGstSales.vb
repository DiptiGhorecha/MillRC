﻿Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmGstSales
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Public dBaseConnection As New System.Data.OleDb.OleDbConnection
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim fnumm As Integer
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim formloaded As Boolean = False

    Private Sub FrmGstSales_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        If DateTime.Now.Month = 1 Then
            ComboBox3.Text = DateAndTime.MonthName(12)
            ComboBox4.Text = DateTime.Now.Year - 1
        Else
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox4.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 1).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            ' TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3))
        End If
        TextBox1.Focus()
        HSNRadio2.Checked = True
        B2BRadio3.Checked = True
        formloaded = True
    End Sub
    Private Sub KeyUpHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.F1 Then
        End If
    End Sub

    Private Sub KeyDownHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        TxtSrch.Text = ""
        If e.ColumnIndex = 1 Then
            indexorder = "[BILL].GROUP"
            GroupBox5.Text = "Search by Group"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 2 Then
            indexorder = "[BILL].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 16 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
    End Sub
    Private Sub ShowData(mnth As String, yr As String)
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME,[PARTY].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
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
            DataGridView2.Columns(15).Visible = False
            DataGridView2.Columns(13).Visible = False
            DataGridView2.Columns(4).Width = 80
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Invoice number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Invoice number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Invoice number must be less than To invoice number")
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
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No. to wsrite to csv file '''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        Dim srno As Integer = 0
        'xpage = 1
        xpage = Val("2")

        Dim numRec As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\GST_sales.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        B2BRadio3.Checked = True
        If ComboBox1.Text = "All" Then
            chkrs1.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = False
        Else
            chkrs1.Open("SELECT * FROM GST where GST_DESC='" & ComboBox1.Text & "' ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = True
        End If
        Dim HSN As String
        Dim HSNPRINT As Boolean
        Dim totnet, tottaxable, totcgst, totsgst, groupnet, grouptaxable, groupcgst, groupsgst As Double
        Dim atotnet, atottaxable, atotcgst, atotsgst, agroupnet, agrouptaxable, agroupcgst, agroupsgst As Double
        Do While chkrs1.EOF = False
            HSN = chkrs1.Fields(0).Value
            HSNPRINT = True
            For X As Integer = strbill To edbill
                If X < DataGridView2.RowCount Then
                    If xcount = 0 Then
                    globalHeader("GST - Sales Report " & ComboBox3.Text & " - " & ComboBox4.Text, fnum, fnumm)
                    Print(fnum, GetStringToPrint(13, "Date   ", "S") & GetStringToPrint(20, "Invoice No", "S") & GetStringToPrint(55, "Party Name", "S") & GetStringToPrint(15, "GST No", "S") & GetStringToPrint(13, "State Code", "S") & GetStringToPrint(13, "State", "S") & GetStringToPrint(15, "HSN /SAC ", "S") & GetStringToPrint(13, "Type", "S") & GetStringToPrint(12, "    Tax Rate", "S") & GetStringToPrint(12, " Quantity", "S") & GetStringToPrint(12, " Basic", "N") & GetStringToPrint(12, "        CGST ", "N") & GetStringToPrint(12, "     SGST ", "N") & GetStringToPrint(12, "     IGST ", "N") & GetStringToPrint(15, "     Total", "N") & GetStringToPrint(15, "    Month", "S") & GetStringToPrint(25, "   Remark", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(13, "Date   ", "S") & "," & GetStringToPrint(20, "Invoice No", "S") & "," & GetStringToPrint(55, "Party Name", "S") & "," & GetStringToPrint(15, "GST No", "S") & "," & GetStringToPrint(13, "State Code", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(15, "HSN /SAC ", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(12, "    Tax Rate", "S") & "," & GetStringToPrint(12, " Quantity", "S") & "," & GetStringToPrint(12, " Basic", "N") & "," & GetStringToPrint(12, "        CGST ", "N") & "," & GetStringToPrint(12, "     SGST ", "N") & "," & GetStringToPrint(12, "     IGST ", "N") & "," & GetStringToPrint(15, "     Total", "N") & "," & GetStringToPrint(15, "    Month", "S") & "," & GetStringToPrint(25, "   Remark", "S") & "," & vbNewLine)

                    Print(fnum, " " & vbNewLine)
                    Print(fnumm, " " & vbNewLine)
                    xcount = xcount + 3
                End If
                Dim partyGST As String = ""
                If HSNRadio1.Checked = True Then
                    If DataGridView2.Item(11, X).Value = HSN Then

                        If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                            partyGST = ""
                        Else
                            partyGST = DataGridView2.Item(17, X).Value
                        End If
                        If B2BRadio1.Checked = True Then
                            If partyGST.Trim <> "" Then
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    ' Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    'Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                    ' xcount = xcount + 1
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                    'Print(fnumm, GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(1, X).Value, "S") & "," & GetStringToPrint(15, DataGridView2.Item(2, X).Value, "S") & "," & GetStringToPrint(15, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(6, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(8, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(17, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    atotnet = atotnet + DataGridView2.Item(10, X).Value
                                    atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                    atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                    agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                    atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                    agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                    agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                    agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    totcgst = totcgst + DataGridView2.Item(7, X).Value
                                    totsgst = totsgst + DataGridView2.Item(9, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                    groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If B2BRadio2.Checked = True Then
                                If partyGST.Trim.Equals("") Then
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        'Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        'Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        'xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                        atotnet = atotnet + DataGridView2.Item(10, X).Value
                                        atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                        atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                        agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                        atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                        agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                        agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                        agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    Else
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        totcgst = totcgst + DataGridView2.Item(7, X).Value
                                        totsgst = totsgst + DataGridView2.Item(9, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                        groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                        groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    End If
                                End If
                            Else
                                If HSNRadio1.Checked = True And HSNPRINT = True Then
                                    ' Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    ' Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                    HSNPRINT = False
                                    ' xcount = xcount + 1
                                End If
                                If DataGridView2.Item(15, X).Value = True Then
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                    atotnet = atotnet + DataGridView2.Item(10, X).Value
                                    atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                    atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                    agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                    atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                    agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                    agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                    agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    totcgst = totcgst + DataGridView2.Item(7, X).Value
                                    totsgst = totsgst + DataGridView2.Item(9, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                    groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        End If
                        '  xcount = xcount + 1
                    End If
                Else


                    If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                        partyGST = ""
                    Else
                        partyGST = DataGridView2.Item(17, X).Value
                    End If
                    If B2BRadio1.Checked = True Then
                        If partyGST.Trim <> "" Then
                            If DataGridView2.Item(15, X).Value = True Then
                                Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                atotnet = atotnet + DataGridView2.Item(10, X).Value
                                atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                xcount = xcount + 1
                            Else
                                Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                totcgst = totcgst + DataGridView2.Item(7, X).Value
                                totsgst = totsgst + DataGridView2.Item(9, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                xcount = xcount + 1
                            End If
                        End If
                    Else
                        If B2BRadio2.Checked = True Then
                            If partyGST.Trim.Equals("") Then
                                If DataGridView2.Item(15, X).Value = True Then
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                    atotnet = atotnet + DataGridView2.Item(10, X).Value
                                    atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                    atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                    agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                    atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                    agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                    agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                    agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    totcgst = totcgst + DataGridView2.Item(7, X).Value
                                    totsgst = totsgst + DataGridView2.Item(9, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                    groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If DataGridView2.Item(15, X).Value = True Then
                                Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                atotnet = atotnet + DataGridView2.Item(10, X).Value
                                atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                xcount = xcount + 1
                            Else
                                Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                totnet = totnet + DataGridView2.Item(10, X).Value
                                totcgst = totcgst + DataGridView2.Item(7, X).Value
                                totsgst = totsgst + DataGridView2.Item(9, X).Value
                                groupnet = groupnet + DataGridView2.Item(10, X).Value
                                tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                xcount = xcount + 1
                            End If
                        End If
                    End If


                End If
                End If
            Next
            If HSNRadio1.Checked = True Then
                If chkrs1.EOF = False Then
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
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(20, "Total with Advance >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(tottaxable + atottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(totcgst + atotcgst, "######0.00"), "N") & GetStringToPrint(12, Format(totsgst + atotsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(totnet + atotnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total with Advance >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(tottaxable + atottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totcgst + atotcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totsgst + atotsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(totnet + atotnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(20, "Total Advance >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(atottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(atotcgst, "######0.00"), "N") & GetStringToPrint(12, Format(atotsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(atotnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total Advance >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(atottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(atotcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(atotsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(atotnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(20, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(tottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(totcgst, "######0.00"), "N") & GetStringToPrint(12, Format(totsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(totnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(tottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(totnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)

        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
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
            '  DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            ' Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim endP As DateTime = startP
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND billwr.tax > 0 order by BILLWR.BILL_DATE,billwr.bill_no;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        groupcgst = 0
        groupsgst = 0
        Dim taxrate As Double = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            '  Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnum, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(55, DataGridView1.Item(12, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(18, X).Value), "", DataGridView1.Item(18, X).Value), "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, " Gujarat", "S") & GetStringToPrint(15, " 997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format(18, "######0.00"), "N") & GetStringToPrint(12, "   ", "S") & GetStringToPrint(12, Format(DataGridView1.Item(4, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(15, Format(DataGridView1.Item(7, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(55, DataGridView1.Item(12, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(18, X).Value), "", DataGridView1.Item(18, X).Value), "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, " Gujarat", "S") & "," & GetStringToPrint(15, " 997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format(18, "######0.00"), "N") & "," & GetStringToPrint(12, "   ", "S") & "," & GetStringToPrint(12, Format(DataGridView1.Item(4, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & "," & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(15, Format(DataGridView1.Item(7, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & GetStringToPrint(25, " ", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & "," & GetStringToPrint(35, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(4, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(12, X).Value, "S") & vbNewLine)
            groupnet = groupnet + DataGridView1.Item(7, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(4, X).Value
            groupcgst = groupcgst + Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2
            groupsgst = groupsgst + Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        ' Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        ' Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)

        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(20, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

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
            '  DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            ' Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim endP As DateTime = startP
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND bill.tax > 0 order by BILL.BILL_DATE,bill.bill_no;")     '(ACCMST.GST_NO<>'' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        groupcgst = 0
        groupsgst = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)

            '   Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)

            '   Print(fnum, GetStringToPrint(16, DataGridView1.Item(38, X).Value, "S") & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            '   Print(fnumm, GetStringToPrint(16, DataGridView1.Item(38, X).Value, "S") & "," & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)

            Print(fnum, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(38, X).Value), "", DataGridView1.Item(38, X).Value), "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, " Gujarat", "S") & GetStringToPrint(15, " 997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format(18, "######0.00"), "N") & GetStringToPrint(12, "   ", "S") & GetStringToPrint(12, Format(DataGridView1.Item(19, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView1.Item(30, X).Value, "##########0.00"), "N") & GetStringToPrint(12, Format(DataGridView1.Item(29, X).Value, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(15, Format(DataGridView1.Item(27, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(38, X).Value), "", DataGridView1.Item(38, X).Value), "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, " Gujarat", "S") & "," & GetStringToPrint(15, " 997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format(18, "######0.00"), "N") & "," & GetStringToPrint(12, "   ", "S") & "," & GetStringToPrint(12, Format(DataGridView1.Item(19, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView1.Item(30, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView1.Item(29, X).Value, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(15, Format(DataGridView1.Item(27, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & GetStringToPrint(25, " ", "S") & vbNewLine)


            groupnet = groupnet + DataGridView1.Item(27, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(19, X).Value
            groupcgst = groupcgst + DataGridView1.Item(29, X).Value
            groupsgst = groupsgst + DataGridView1.Item(30, X).Value
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        ' Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        ' Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)

        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(20, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill end


        FileClose(fnum)
        FileClose(fnumm)
        FrmGSTSalesView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\GST_sales.dat", RichTextBoxStreamType.PlainText)
        FrmGSTSalesView.Show()
        MsgBox(Application.StartupPath + " \Reports\" & TextBox5.Text & ".CSV file is generated")
    End Sub
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
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Invoice number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Invoice number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Invoice number must be less than To invoice number")
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
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No. to wsrite to csv file '''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        Dim srno As Integer = 0
        'xpage = 1
        xpage = Val("2")

        Dim numRec As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\GST_sales.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        B2BRadio3.Checked = True
        If ComboBox1.Text = "All" Then
            chkrs1.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = False
        Else
            chkrs1.Open("SELECT * FROM GST where GST_DESC='" & ComboBox1.Text & "' ORDER BY HSN_NO", xcon)
            HSNRadio1.Checked = True
        End If
        Dim HSN As String
        Dim HSNPRINT As Boolean
        Dim totnet, tottaxable, totcgst, totsgst, groupnet, grouptaxable, groupcgst, groupsgst As Double
        Dim atotnet, atottaxable, atotcgst, atotsgst, agroupnet, agrouptaxable, agroupcgst, agroupsgst As Double
        Do While chkrs1.EOF = False
            HSN = chkrs1.Fields(0).Value
            HSNPRINT = True
            For X As Integer = strbill To edbill
                If X < DataGridView2.RowCount Then
                    If xcount = 0 Then
                        globalHeader("GST - Sales Report " & ComboBox3.Text & " - " & ComboBox4.Text, fnum, fnumm)
                        Print(fnum, GetStringToPrint(13, "Date   ", "S") & GetStringToPrint(12, "Invoice No", "S") & GetStringToPrint(55, "Party Name", "S") & GetStringToPrint(15, "GST No", "S") & GetStringToPrint(13, "State Code", "S") & GetStringToPrint(9, "State", "S") & GetStringToPrint(10, "HSN /SAC ", "S") & GetStringToPrint(6, "Type", "S") & GetStringToPrint(12, "    Tax Rate", "S") & GetStringToPrint(10, " Quantity", "S") & GetStringToPrint(12, " Basic", "N") & GetStringToPrint(12, "        CGST ", "N") & GetStringToPrint(12, "     SGST ", "N") & GetStringToPrint(12, "     IGST ", "N") & GetStringToPrint(15, "     Total", "N") & GetStringToPrint(15, "    Month", "S") & GetStringToPrint(25, "   Remark", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(13, "Date   ", "S") & "," & GetStringToPrint(20, "Invoice No", "S") & "," & GetStringToPrint(55, "Party Name", "S") & "," & GetStringToPrint(15, "GST No", "S") & "," & GetStringToPrint(13, "State Code", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(15, "HSN /SAC ", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(12, "    Tax Rate", "S") & "," & GetStringToPrint(12, " Quantity", "S") & "," & GetStringToPrint(12, " Basic", "N") & "," & GetStringToPrint(12, "        CGST ", "N") & "," & GetStringToPrint(12, "     SGST ", "N") & "," & GetStringToPrint(12, "     IGST ", "N") & "," & GetStringToPrint(15, "     Total", "N") & "," & GetStringToPrint(15, "    Month", "S") & "," & GetStringToPrint(25, "   Remark", "S") & "," & vbNewLine)

                        Print(fnum, " " & vbNewLine)
                        Print(fnumm, " " & vbNewLine)
                        xcount = xcount + 3
                    End If
                    Dim partyGST As String = ""
                    If HSNRadio1.Checked = True Then
                        If DataGridView2.Item(11, X).Value = HSN Then

                            If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                                partyGST = ""
                            Else
                                partyGST = DataGridView2.Item(17, X).Value
                            End If
                            If B2BRadio1.Checked = True Then
                                If partyGST.Trim <> "" Then
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        ' Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        'Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        ' xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                        'Print(fnumm, GetStringToPrint(35, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(1, X).Value, "S") & "," & GetStringToPrint(15, DataGridView2.Item(2, X).Value, "S") & "," & GetStringToPrint(15, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(6, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(8, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(17, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value, "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        atotnet = atotnet + DataGridView2.Item(10, X).Value
                                        atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                        atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                        agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                        atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                        agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                        agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                        agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    Else
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        totcgst = totcgst + DataGridView2.Item(7, X).Value
                                        totsgst = totsgst + DataGridView2.Item(9, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                        groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                        groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    End If
                                End If
                            Else
                                If B2BRadio2.Checked = True Then
                                    If partyGST.Trim.Equals("") Then
                                        If HSNRadio1.Checked = True And HSNPRINT = True Then
                                            'Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                            'Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                            HSNPRINT = False
                                            'xcount = xcount + 1
                                        End If
                                        If DataGridView2.Item(15, X).Value = True Then
                                            Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                            Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                            atotnet = atotnet + DataGridView2.Item(10, X).Value
                                            atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                            atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                            agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                            atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                            agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                            agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                            agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                            xcount = xcount + 1
                                        Else
                                            Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                            Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                            totnet = totnet + DataGridView2.Item(10, X).Value
                                            totcgst = totcgst + DataGridView2.Item(7, X).Value
                                            totsgst = totsgst + DataGridView2.Item(9, X).Value
                                            groupnet = groupnet + DataGridView2.Item(10, X).Value
                                            tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                            grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                            groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                            groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                            xcount = xcount + 1
                                        End If
                                    End If
                                Else
                                    If HSNRadio1.Checked = True And HSNPRINT = True Then
                                        ' Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        ' Print(fnumm, GetStringToPrint(35, "HSN Number :" & chkrs1.Fields(0).Value, "S") & "," & GetStringToPrint(75, chkrs1.Fields(1).Value, "S") & vbNewLine)
                                        HSNPRINT = False
                                        ' xcount = xcount + 1
                                    End If
                                    If DataGridView2.Item(15, X).Value = True Then
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                        atotnet = atotnet + DataGridView2.Item(10, X).Value
                                        atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                        atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                        agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                        atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                        agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                        agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                        agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    Else
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        totcgst = totcgst + DataGridView2.Item(7, X).Value
                                        totsgst = totsgst + DataGridView2.Item(9, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                        groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                        groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    End If
                                End If
                            End If
                            '  xcount = xcount + 1
                        End If
                    Else


                        If IsDBNull(DataGridView2.Item(17, X).Value) Or DataGridView2.Item(17, X).Value Is Nothing Then
                            partyGST = ""
                        Else
                            partyGST = DataGridView2.Item(17, X).Value
                        End If
                        If B2BRadio1.Checked = True Then
                            If partyGST.Trim <> "" Then
                                If DataGridView2.Item(15, X).Value = True Then
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                    atotnet = atotnet + DataGridView2.Item(10, X).Value
                                    atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                    atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                    agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                    atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                    agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                    agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                    agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    totcgst = totcgst + DataGridView2.Item(7, X).Value
                                    totsgst = totsgst + DataGridView2.Item(9, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                    groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        Else
                            If B2BRadio2.Checked = True Then
                                If partyGST.Trim.Equals("") Then
                                    If DataGridView2.Item(15, X).Value = True Then
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                        atotnet = atotnet + DataGridView2.Item(10, X).Value
                                        atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                        atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                        agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                        atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                        agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                        agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                        agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    Else
                                        Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                        totnet = totnet + DataGridView2.Item(10, X).Value
                                        totcgst = totcgst + DataGridView2.Item(7, X).Value
                                        totsgst = totsgst + DataGridView2.Item(9, X).Value
                                        groupnet = groupnet + DataGridView2.Item(10, X).Value
                                        tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                        grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                        groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                        groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                        xcount = xcount + 1
                                    End If
                                End If
                            Else
                                If DataGridView2.Item(15, X).Value = True Then
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & GetStringToPrint(25, "Against Advance", "S") & "," & vbNewLine)
                                    atotnet = atotnet + DataGridView2.Item(10, X).Value
                                    atotcgst = atotcgst + DataGridView2.Item(7, X).Value
                                    atotsgst = atotsgst + DataGridView2.Item(9, X).Value
                                    agroupnet = agroupnet + DataGridView2.Item(10, X).Value
                                    atottaxable = atottaxable + DataGridView2.Item(5, X).Value
                                    agrouptaxable = agrouptaxable + DataGridView2.Item(5, X).Value
                                    agroupcgst = agroupcgst + DataGridView2.Item(7, X).Value
                                    agroupsgst = agroupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                Else
                                    Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(12, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, "Gujarat", "S") & GetStringToPrint(10, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, partyGST, "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, "Gujarat", "S") & "," & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & " " & "," & vbNewLine)
                                    totnet = totnet + DataGridView2.Item(10, X).Value
                                    totcgst = totcgst + DataGridView2.Item(7, X).Value
                                    totsgst = totsgst + DataGridView2.Item(9, X).Value
                                    groupnet = groupnet + DataGridView2.Item(10, X).Value
                                    tottaxable = tottaxable + DataGridView2.Item(5, X).Value
                                    grouptaxable = grouptaxable + DataGridView2.Item(5, X).Value
                                    groupcgst = groupcgst + DataGridView2.Item(7, X).Value
                                    groupsgst = groupsgst + DataGridView2.Item(9, X).Value
                                    xcount = xcount + 1
                                End If
                            End If
                        End If


                    End If
                End If
            Next
            If HSNRadio1.Checked = True Then
                If chkrs1.EOF = False Then
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
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, "Total with ", "S") & GetStringToPrint(12, "Advance >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(9, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(tottaxable + atottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(totcgst + atotcgst, "######0.00"), "N") & GetStringToPrint(12, Format(totsgst + atotsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(totnet + atotnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total with Advance >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(tottaxable + atottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totcgst + atotcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totsgst + atotsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(totnet + atotnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, "Total ", "S") & GetStringToPrint(12, "Advance >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(9, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(atottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(atotcgst, "######0.00"), "N") & GetStringToPrint(12, Format(atotsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(atotnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total Advance >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(atottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(atotcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(atotsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(atotnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(12, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(9, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(tottaxable, "######0.00"), "N") & GetStringToPrint(12, Format(totcgst, "######0.00"), "N") & GetStringToPrint(12, Format(totsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(totnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(tottaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(totsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(totnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)

        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
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
            '  DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            ' Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim endP As DateTime = startP
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND billwr.tax > 0 order by BILLWR.BILL_DATE,billwr.bill_no;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        groupcgst = 0
        groupsgst = 0
        Dim taxrate As Double = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            '  Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnum, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(12, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(55, DataGridView1.Item(12, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(18, X).Value), "", DataGridView1.Item(18, X).Value), "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, " Gujarat", "S") & GetStringToPrint(10, " 997212", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format(18, "######0.00"), "N") & GetStringToPrint(10, "   ", "S") & GetStringToPrint(12, Format(DataGridView1.Item(4, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(15, Format(DataGridView1.Item(7, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(55, DataGridView1.Item(12, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(18, X).Value), "", DataGridView1.Item(18, X).Value), "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, " Gujarat", "S") & "," & GetStringToPrint(15, " 997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format(18, "######0.00"), "N") & "," & GetStringToPrint(12, "   ", "S") & "," & GetStringToPrint(12, Format(DataGridView1.Item(4, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & "," & GetStringToPrint(12, Format(Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2, "##########0.00"), "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(15, Format(DataGridView1.Item(7, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & GetStringToPrint(25, " ", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & "," & GetStringToPrint(35, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(7, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(4, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(12, X).Value, "S") & vbNewLine)
            groupnet = groupnet + DataGridView1.Item(7, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(4, X).Value
            groupcgst = groupcgst + Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2
            groupsgst = groupsgst + Convert.ToDouble(DataGridView1.Item(6, X).Value) / 2
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        ' Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        ' Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)

        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(12, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(9, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

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
            '  DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            ' Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox5.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            Dim endP As DateTime = startP
            Dim CurrD As DateTime = startP

            ''   packtable()


            Dim tb As DataTable
            'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
            tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND bill.tax > 0 order by BILL.BILL_DATE,bill.bill_no;")     '(ACCMST.GST_NO<>'' OR ACCMST.GST_NO='Null') ;")
            DataGridView1.DataSource = tb
            '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        groupnet = 0
        grouptaxable = 0
        groupcgst = 0
        groupsgst = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)

            '   Print(fnum, GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & GetStringToPrint(20, "GO-" & DataGridView2.Item(0, X).Value, "S") & GetStringToPrint(55, DataGridView2.Item(16, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, partyGST, "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(13, "Gujarat", "S") & GetStringToPrint(15, DataGridView2.Item(11, X).Value, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format((DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value), "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(DataGridView2.Item(5, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView2.Item(9, X).Value, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(DataGridView2.Item(10, X).Value, "###########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, "Against Advance", "S") & vbNewLine)

            '   Print(fnum, GetStringToPrint(16, DataGridView1.Item(38, X).Value, "S") & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            '   Print(fnumm, GetStringToPrint(16, DataGridView1.Item(38, X).Value, "S") & "," & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)

            Print(fnum, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(12, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value.ToString.Replace(",", ""), "S") & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(38, X).Value), "", DataGridView1.Item(38, X).Value), "S") & GetStringToPrint(13, "   24", "S") & GetStringToPrint(9, " Gujarat", "S") & GetStringToPrint(10, " 997212", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, Format(18, "######0.00"), "N") & GetStringToPrint(10, "   ", "S") & GetStringToPrint(12, Format(DataGridView1.Item(19, X).Value, "######0.00"), "N") & GetStringToPrint(12, Format(DataGridView1.Item(30, X).Value, "##########0.00"), "N") & GetStringToPrint(12, Format(DataGridView1.Item(29, X).Value, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(15, Format(DataGridView1.Item(27, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & " " & GetStringToPrint(25, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(20, "WF-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value.ToString.Replace(",", ""), "S") & "," & GetStringToPrint(15, IIf(IsDBNull(DataGridView1.Item(38, X).Value), "", DataGridView1.Item(38, X).Value), "S") & "," & GetStringToPrint(13, "   24", "S") & "," & GetStringToPrint(13, " Gujarat", "S") & "," & GetStringToPrint(15, " 997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, Format(18, "######0.00"), "N") & "," & GetStringToPrint(12, "   ", "S") & "," & GetStringToPrint(12, Format(DataGridView1.Item(19, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView1.Item(30, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(12, Format(DataGridView1.Item(29, X).Value, "##########0.00"), "N") & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(15, Format(DataGridView1.Item(27, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " " & ComboBox3.Text & "-" & ComboBox4.Text, "S") & "," & GetStringToPrint(25, " ", "S") & vbNewLine)


            groupnet = groupnet + DataGridView1.Item(27, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(19, X).Value
            groupcgst = groupcgst + DataGridView1.Item(29, X).Value
            groupsgst = groupsgst + DataGridView1.Item(30, X).Value
        Next
        Print(fnumm, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        ' Print(fnum, GetStringToPrint(16, "Total --> ", "S") & GetStringToPrint(35, DataGridView1.RowCount, "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, " ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & " " & GetStringToPrint(55, " ", "S") & vbNewLine)
        ' Print(fnumm, GetStringToPrint(16, "Total --> ", "S") & "," & GetStringToPrint(35, DataGridView1.RowCount, "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(10, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)

        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(12, "Total >", "S") & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(9, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(6, " ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(10, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(25, " ", "S") & vbNewLine)
        Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(20, "Total >", "S") & "," & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(25, " ", "S") & "," & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)

        ''''''''''''''''''''''''''''''''''''''''''''''getting data from dbase files of whbill end

        FileClose(fnum)
        FileClose(fnumm)
        FrmGSTSalesView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\GST_sales.dat", RichTextBoxStreamType.PlainText)
        FrmGSTSalesView.Show()
        CreatePDF(Application.StartupPath & "\Reports\GST_sales.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True
        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 10
            pdfPage.Width = 842
            pdfPage.Height = 595
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 5, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1
                line = readFile.ReadLine()
                If counter >= 43 Then
                    counter = 0
                    pdfPage = pdf.AddPage()
                    graph = XGraphics.FromPdfPage(pdfPage)
                    font = New XFont("COURIER NEW", 5, XFontStyle.Regular)

                    pdfPage.TrimMargins.Left = 10

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

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 1).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            ' TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3))
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
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 1).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            'TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3))
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub


    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        Dim i As Integer = DataGridView2.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(0, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView2.Item(0, i).Value)
            TextBox4.Text = DataGridView2.CurrentCell.RowIndex
        End If
        ' GroupBox5.Visible = False
        ' DataGridView2.Visible = False
        ' Me.Width = Me.Width - DataGridView2.Width + 15
        ' Me.Height = Me.Height - 145
        If ctrlname = "TextBox1" Then
            TextBox3.Text = DataGridView2.CurrentCell.RowIndex
            TextBox2.Focus()
        Else
            If ctrlname = "TextBox2" Then
                TextBox4.Text = DataGridView2.CurrentCell.RowIndex
                Button1.Focus()
            Else
                TextBox1.Focus()
            End If
        End If
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        Console.WriteLine(TxtSrch.Text)
        da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME,[PARTY].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR([BILL].BILL_DATE)='" & ComboBox4.Text & "' and " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [BILL].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        ctrlname = "TextBox1"
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        ctrlname = "TextBox2"
    End Sub

    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        Dim i As Integer = DataGridView2.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(0, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView2.Item(0, i).Value)
            TextBox4.Text = DataGridView2.CurrentCell.RowIndex
        End If
        ' GroupBox5.Visible = False
        ' DataGridView2.Visible = False
        ' Me.Width = Me.Width - DataGridView2.Width + 15
        ' Me.Height = Me.Height - 145
        If ctrlname = "TextBox1" Then
            TextBox3.Text = DataGridView2.CurrentCell.RowIndex
            TextBox2.Focus()
        Else
            If ctrlname = "TextBox2" Then
                TextBox4.Text = DataGridView2.CurrentCell.RowIndex
                Button1.Focus()
            Else
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub FrmGstSales_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub FrmGstSales_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "TextBox1" Or Me.ActiveControl.Name = "TextBox2") Then
            DataGridView2.Visible = True
            GroupBox5.Visible = True
            Me.Width = Me.Width + DataGridView2.Width - 15
            Me.Height = Me.Height + 145
            ctrlname = Me.ActiveControl.Name
            ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 2 Then
            B2BRadio1.Enabled = True
            B2BRadio2.Enabled = True
        Else
            B2BRadio1.Enabled = False
            B2BRadio2.Enabled = False
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Trim.Equals("") Then
        Else
            For i As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Rows(i).Cells(0).Value IsNot Nothing Then
                    '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
                    If Convert.ToInt32(DataGridView2.Rows(i).Cells(0).Value) > (Convert.ToInt32(TextBox1.Text)) Then
                        DataGridView2.ClearSelection()
                        If i = 0 Then
                            DataGridView2.Rows(i).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                        Else
                            DataGridView2.Rows(i - 1).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
                        End If
                        'DataGridView2.Rows(i).Cells(0).Selected = True
                        'DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        ''DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        'TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1  'DataGridView2.Rows(i).Cells(0).Value   'DataGridView2.SelectedCells.Item(0).Value
                        'MessageBox.Show(irowindex)
                        Exit For
                    Else
                        DataGridView2.Rows(i).Cells(0).Selected = True
                        DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Trim.Equals("") Then
        Else
            For i As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Rows(i).Cells(0).Value IsNot Nothing Then
                    '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
                    If Convert.ToInt32(DataGridView2.Rows(i).Cells(0).Value) > (Convert.ToInt32(TextBox2.Text)) Then
                        DataGridView2.ClearSelection()
                        If i = 0 Then
                            DataGridView2.Rows(i).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                        Else
                            DataGridView2.Rows(i - 1).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
                        End If
                        'DataGridView2.Rows(i).Cells(0).Selected = True
                        'DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        ''DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        'TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1  'DataGridView2.Rows(i).Cells(0).Value   'DataGridView2.SelectedCells.Item(0).Value
                        'MessageBox.Show(irowindex)
                        Exit For
                    Else
                        DataGridView2.Rows(i).Cells(0).Selected = True
                        DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ComboBox4_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox4.TextUpdate
        If ComboBox4.FindString(ComboBox4.Text) < 0 Then
            ComboBox4.Text = ComboBox4.Text.Remove(ComboBox4.Text.Length - 1)
            ComboBox4.SelectionStart = ComboBox4.Text.Length
        End If
    End Sub
    Private Sub ComboBox3_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox3.TextUpdate
        If ComboBox3.FindString(ComboBox3.Text) < 0 Then
            ComboBox3.Text = ComboBox3.Text.Remove(ComboBox3.Text.Length - 1)
            ComboBox3.SelectionStart = ComboBox3.Text.Length
        End If
    End Sub

End Class