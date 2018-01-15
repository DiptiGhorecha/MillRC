Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmInvoiceSummary
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"

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

    Private Sub FrmInvoiceSummary_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        ' GroupBox5.Visible = False
        '  DataGridView2.Visible = False
        If DateTime.Now.Month = 1 Then
            ComboBox3.Text = DateAndTime.MonthName(12)
            ComboBox4.Text = DateTime.Now.Year - 1
            ComboBox2.Text = DateAndTime.MonthName(12)
            ComboBox1.Text = DateTime.Now.Year - 1
        Else
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox4.Text = DateTime.Now.Year
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox1.Text = DateTime.Now.Year
        End If
        'ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        'ComboBox4.Text = DateTime.Now.Year
        'ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        'ComboBox1.Text = DateTime.Now.Year
        'ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        'TextBox1.Text = DataGridView2.Item(0, 0).Value
        'TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
        'TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
        'TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
        'TextBox1.Focus()
        'HSNRadio2.Checked = True
        ' B2BRadio3.Checked = True
        formloaded = True
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR
        Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '  MsgBox(DaysInMonth)
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        Dim CurrD As DateTime = startP
        '  MsgBox(startP)
        '  MsgBox(endP)



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
                FileOpen(fnum, Application.StartupPath & "\Invoicessummary.dat", OpenMode.Output)
                FileOpen(fnumm, Application.StartupPath & "\" & TextBox5.Text & ".csv", OpenMode.Output)
                If xcon.State = ConnectionState.Open Then
                Else
                    xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                End If

        Dim totnet, tottaxable, totcgst, totsgst, groupnet, grouptaxable, groupcgst, groupsgst, b2btaxable, b2ctaxable, ntaxable As Double
        Dim mtotnet, mtottaxable, mtotcgst, mtotsgst, mgroupnet, mgrouptaxable, mgroupcgst, mgroupsgst, mb2btaxable, mb2ctaxable, mntaxable As Double
        mtotnet = mtottaxable = mtotcgst = mtotsgst = mgroupnet = mgrouptaxable = mgroupcgst = mgroupsgst = mb2btaxable = mb2ctaxable = mntaxable = 0
        While (CurrD <= endP)
            ' ProcessData(CurrD)
            'Console.WriteLine(CurrD.ToShortDateString)
            mtotnet = 0
            mtottaxable = 0
            mtotcgst = 0
            mtotsgst = 0
            mgroupnet = 0
            mgrouptaxable = 0
            mgroupcgst = 0
            mgroupsgst = 0
            mb2btaxable = 0
            mb2ctaxable = 0
            mntaxable = 0
            If xcount = 0 Then
                Print(fnum, GetStringToPrint(15, "Month-Year", "S") & GetStringToPrint(30, "        Total Rent Residential", "S") & GetStringToPrint(30, "    Total Rent Non-Residential", "S") & GetStringToPrint(13, "         CGST", "S") & GetStringToPrint(13, "         SGST", "S") & GetStringToPrint(15, "      Net Total", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(15, "Month-Year", "S") & "," & GetStringToPrint(30, "        Total Rent Residential", "S") & "," & GetStringToPrint(30, "    Total Rent Non-Residential", "S") & "," & GetStringToPrint(13, "         CGST", "S") & "," & GetStringToPrint(13, "         SGST", "S") & "," & GetStringToPrint(15, "      Net Total", "S") & vbNewLine)
                Print(fnum, " " & vbNewLine)
                Print(fnumm, " " & vbNewLine)
                xcount = xcount + 3
            End If
            Dim str As String = "SELECT [BILL].*,[PARTY].P_NAME,[party].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE WHERE [BILL].BILL_DATE=#" & CurrD & "# order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO"

            chkrs1.Open("SELECT [BILL].*,[PARTY].P_NAME,[party].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE WHERE [BILL].BILL_DATE=#" & CurrD & "# order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", xcon)
            Do While chkrs1.EOF = False
                Dim partyGST As String = ""
                If chkrs1.Fields(11).Value.Equals("997212") Then
                    partyGST = ""
                    If IsDBNull(chkrs1.Fields(17).Value) Or chkrs1.Fields(17).Value Is Nothing Then
                        partyGST = ""
                        If chkrs1.Fields(15).Value = True Then
                        Else
                            b2ctaxable = b2ctaxable + chkrs1.Fields(5).Value
                            mb2ctaxable = mb2ctaxable + chkrs1.Fields(5).Value
                        End If
                    Else
                        If chkrs1.Fields(15).Value = True Then
                        Else
                            b2btaxable = b2btaxable + chkrs1.Fields(5).Value
                            mb2btaxable = mb2btaxable + chkrs1.Fields(5).Value
                        End If
                    End If
                Else
                    If chkrs1.Fields(15).Value = True Then
                    Else
                        ntaxable = ntaxable + chkrs1.Fields(5).Value
                        mntaxable = mntaxable + chkrs1.Fields(5).Value
                    End If
                End If
                If chkrs1.Fields(15).Value = True Then
                Else
                    totnet = totnet + chkrs1.Fields(10).Value
                    totcgst = totcgst + chkrs1.Fields(7).Value
                    totsgst = totsgst + chkrs1.Fields(9).Value
                    groupnet = groupnet + chkrs1.Fields(10).Value
                    tottaxable = tottaxable + chkrs1.Fields(5).Value
                    grouptaxable = grouptaxable + chkrs1.Fields(5).Value
                    groupcgst = groupcgst + chkrs1.Fields(7).Value
                    groupsgst = groupsgst + chkrs1.Fields(9).Value

                    mtotnet = mtotnet + chkrs1.Fields(10).Value
                    mtotcgst = mtotcgst + chkrs1.Fields(7).Value
                    mtotsgst = mtotsgst + chkrs1.Fields(9).Value
                    mgroupnet = mgroupnet + chkrs1.Fields(10).Value
                    mtottaxable = mtottaxable + chkrs1.Fields(5).Value
                    mgrouptaxable = mgrouptaxable + chkrs1.Fields(5).Value
                    mgroupcgst = mgroupcgst + chkrs1.Fields(7).Value
                    mgroupsgst = mgroupsgst + chkrs1.Fields(9).Value

                    xcount = xcount + 1
                End If

                If chkrs1.EOF = True Then
                    Exit Do
                Else
                    chkrs1.MoveNext()
                End If

            Loop
            chkrs1.Close()
            Print(fnum, GetStringToPrint(15, CurrD.Month & "-" & CurrD.Year, "S") & GetStringToPrint(30, Format(mntaxable, "##########0.00"), "N") & GetStringToPrint(30, Format(mb2btaxable + mb2ctaxable, "##########0.00"), "N") & GetStringToPrint(13, Format(mtotcgst, "######0.00"), "N") & GetStringToPrint(13, Format(mtotsgst, "##########0.00"), "N") & GetStringToPrint(15, Format(mtotnet, "##########0.00"), "N") & vbNewLine)
            Print(fnumm, GetStringToPrint(15, CurrD.Month & "-" & CurrD.Year, "S") & "," & GetStringToPrint(30, Format(mntaxable, "##########0.00"), "N") & "," & GetStringToPrint(30, Format(mb2btaxable + mb2ctaxable, "##########0.00"), "N") & "," & GetStringToPrint(13, Format(mtotcgst, "######0.00"), "N") & "," & GetStringToPrint(13, Format(mtotsgst, "##########0.00"), "N") & "," & GetStringToPrint(15, Format(mtotnet, "##########0.00"), "N") & vbNewLine)
            '  Print(fnumm, GetStringToPrint(35, CurrD.Month & "-" & CurrD.Year, "S") & "," & GetStringToPrint(35, DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)

            CurrD = CurrD.AddMonths(1)
            mtotnet = mtottaxable = mtotcgst = mtotsgst = mgroupnet = mgrouptaxable = mgroupcgst = mgroupsgst = mb2btaxable = mb2ctaxable = mntaxable = 0
        End While
        'If HSNRadio1.Checked = True Then
        '            If chkrs1.EOF = False Then
        '                If groupnet > 0 Then
        '                    Print(fnum, " " & vbNewLine)
        '                    Print(fnumm, " " & vbNewLine)
        '                    Print(fnum, GetStringToPrint(35, "Group Total --> ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, "  ", "S") & " " & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & vbNewLine)
        '                    Print(fnumm, GetStringToPrint(35, "Group Total --> ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(17, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(15, "  ", "S") & ", " & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
        '                    groupnet = 0
        '                    grouptaxable = 0
        '                    groupcgst = 0
        '                    groupsgst = 0
        '                End If
        '                chkrs1.MoveNext()
        '            End If
        '            If chkrs1.EOF = True Then
        '                Exit Do
        '            End If
        '        Else
        '            Exit Do
        '        End If
        '    Loop
        '    chkrs1.Close()
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(15, "Total -->", "S") & GetStringToPrint(30, Format(ntaxable, "##########0.00"), "N") & GetStringToPrint(30, Format(b2btaxable + b2ctaxable, "##########0.00"), "N") & GetStringToPrint(13, Format(totcgst, "######0.00"), "N") & GetStringToPrint(13, Format(totsgst, "##########0.00"), "N") & GetStringToPrint(15, Format(totnet, "##########0.00"), "N") & vbNewLine)
        Print(fnumm, GetStringToPrint(15, "Total -->", "S") & "," & GetStringToPrint(30, Format(ntaxable, "##########0.00"), "N") & "," & GetStringToPrint(30, Format(b2btaxable + b2ctaxable, "##########0.00"), "N") & "," & GetStringToPrint(13, Format(totcgst, "######0.00"), "N") & "," & GetStringToPrint(13, Format(totsgst, "##########0.00"), "N") & "," & GetStringToPrint(15, Format(totnet, "##########0.00"), "N") & vbNewLine)
        FileClose(fnum)
        FileClose(fnumm)
        Form14.RichTextBox1.LoadFile(Application.StartupPath & "\invoicessummary.dat", RichTextBoxStreamType.PlainText)
        Form14.Show()
        MsgBox(Application.StartupPath + " \" & TextBox5.Text & ".CSV file is generated")


    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR
        Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '  MsgBox(DaysInMonth)
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        Dim CurrD As DateTime = startP
        '  MsgBox(startP)
        '  MsgBox(endP)



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
        FileOpen(fnum, Application.StartupPath & "\Invoicessummary.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If

        Dim totnet, tottaxable, totcgst, totsgst, groupnet, grouptaxable, groupcgst, groupsgst, b2btaxable, b2ctaxable, ntaxable As Double
        Dim mtotnet, mtottaxable, mtotcgst, mtotsgst, mgroupnet, mgrouptaxable, mgroupcgst, mgroupsgst, mb2btaxable, mb2ctaxable, mntaxable As Double
        mtotnet = mtottaxable = mtotcgst = mtotsgst = mgroupnet = mgrouptaxable = mgroupcgst = mgroupsgst = mb2btaxable = mb2ctaxable = mntaxable = 0
        While (CurrD <= endP)
            ' ProcessData(CurrD)
            'Console.WriteLine(CurrD.ToShortDateString)
            mtotnet = 0
            mtottaxable = 0
            mtotcgst = 0
            mtotsgst = 0
            mgroupnet = 0
            mgrouptaxable = 0
            mgroupcgst = 0
            mgroupsgst = 0
            mb2btaxable = 0
            mb2ctaxable = 0
            mntaxable = 0
            If xcount = 0 Then
                Print(fnum, GetStringToPrint(15, "Month-Year", "S") & GetStringToPrint(30, "        Total Rent Residential", "S") & GetStringToPrint(30, "    Total Rent Non-Residential", "S") & GetStringToPrint(13, "         CGST", "S") & GetStringToPrint(13, "         SGST", "S") & GetStringToPrint(15, "      Net Total", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(15, "Month-Year", "S") & "," & GetStringToPrint(30, "        Total Rent Residential", "S") & "," & GetStringToPrint(30, "    Total Rent Non-Residential", "S") & "," & GetStringToPrint(13, "         CGST", "S") & "," & GetStringToPrint(13, "         SGST", "S") & "," & GetStringToPrint(15, "      Net Total", "S") & vbNewLine)
                Print(fnum, " " & vbNewLine)
                Print(fnumm, " " & vbNewLine)
                xcount = xcount + 3
            End If
            Dim str As String = "SELECT [BILL].*,[PARTY].P_NAME,[party].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE WHERE [BILL].BILL_DATE=#" & CurrD & "# order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO"

            chkrs1.Open("SELECT [BILL].*,[PARTY].P_NAME,[party].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE WHERE [BILL].BILL_DATE=#" & CurrD & "# order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", xcon)
            Do While chkrs1.EOF = False
                Dim partyGST As String = ""
                If chkrs1.Fields(11).Value.Equals("997212") Then
                    partyGST = ""
                    If IsDBNull(chkrs1.Fields(17).Value) Or chkrs1.Fields(17).Value Is Nothing Then
                        partyGST = ""
                        If chkrs1.Fields(15).Value = True Then
                        Else
                            b2ctaxable = b2ctaxable + chkrs1.Fields(5).Value
                            mb2ctaxable = mb2ctaxable + chkrs1.Fields(5).Value
                        End If
                    Else
                        If chkrs1.Fields(15).Value = True Then
                        Else
                            b2btaxable = b2btaxable + chkrs1.Fields(5).Value
                            mb2btaxable = mb2btaxable + chkrs1.Fields(5).Value
                        End If
                    End If
                Else
                    If chkrs1.Fields(15).Value = True Then
                    Else
                        ntaxable = ntaxable + chkrs1.Fields(5).Value
                        mntaxable = mntaxable + chkrs1.Fields(5).Value
                    End If
                End If
                If chkrs1.Fields(15).Value = True Then
                Else
                    totnet = totnet + chkrs1.Fields(10).Value
                    totcgst = totcgst + chkrs1.Fields(7).Value
                    totsgst = totsgst + chkrs1.Fields(9).Value
                    groupnet = groupnet + chkrs1.Fields(10).Value
                    tottaxable = tottaxable + chkrs1.Fields(5).Value
                    grouptaxable = grouptaxable + chkrs1.Fields(5).Value
                    groupcgst = groupcgst + chkrs1.Fields(7).Value
                    groupsgst = groupsgst + chkrs1.Fields(9).Value

                    mtotnet = mtotnet + chkrs1.Fields(10).Value
                    mtotcgst = mtotcgst + chkrs1.Fields(7).Value
                    mtotsgst = mtotsgst + chkrs1.Fields(9).Value
                    mgroupnet = mgroupnet + chkrs1.Fields(10).Value
                    mtottaxable = mtottaxable + chkrs1.Fields(5).Value
                    mgrouptaxable = mgrouptaxable + chkrs1.Fields(5).Value
                    mgroupcgst = mgroupcgst + chkrs1.Fields(7).Value
                    mgroupsgst = mgroupsgst + chkrs1.Fields(9).Value

                    xcount = xcount + 1
                End If

                If chkrs1.EOF = True Then
                    Exit Do
                Else
                    chkrs1.MoveNext()
                End If

            Loop
            chkrs1.Close()
            Print(fnum, GetStringToPrint(15, CurrD.Month & "-" & CurrD.Year, "S") & GetStringToPrint(30, Format(mntaxable, "##########0.00"), "N") & GetStringToPrint(30, Format(mb2btaxable + mb2ctaxable, "##########0.00"), "N") & GetStringToPrint(13, Format(mtotcgst, "######0.00"), "N") & GetStringToPrint(13, Format(mtotsgst, "##########0.00"), "N") & GetStringToPrint(15, Format(mtotnet, "##########0.00"), "N") & vbNewLine)
            Print(fnumm, GetStringToPrint(15, CurrD.Month & "-" & CurrD.Year, "S") & "," & GetStringToPrint(30, Format(mntaxable, "##########0.00"), "N") & "," & GetStringToPrint(30, Format(mb2btaxable + mb2ctaxable, "##########0.00"), "N") & "," & GetStringToPrint(13, Format(mtotcgst, "######0.00"), "N") & "," & GetStringToPrint(13, Format(mtotsgst, "##########0.00"), "N") & "," & GetStringToPrint(15, Format(mtotnet, "##########0.00"), "N") & vbNewLine)
            '  Print(fnumm, GetStringToPrint(35, CurrD.Month & "-" & CurrD.Year, "S") & "," & GetStringToPrint(35, DataGridView2.Item(0, X).Value, "S") & "," & GetStringToPrint(13, DataGridView2.Item(4, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView2.Item(10, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(DataGridView2.Item(6, X).Value + DataGridView2.Item(8, X).Value, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView2.Item(5, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView2.Item(15, X).Value, "S") & vbNewLine)

            CurrD = CurrD.AddMonths(1)
            mtotnet = mtottaxable = mtotcgst = mtotsgst = mgroupnet = mgrouptaxable = mgroupcgst = mgroupsgst = mb2btaxable = mb2ctaxable = mntaxable = 0
        End While
        'If HSNRadio1.Checked = True Then
        '            If chkrs1.EOF = False Then
        '                If groupnet > 0 Then
        '                    Print(fnum, " " & vbNewLine)
        '                    Print(fnumm, " " & vbNewLine)
        '                    Print(fnum, GetStringToPrint(35, "Group Total --> ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & GetStringToPrint(12, " ", "S") & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & GetStringToPrint(15, Format(groupnet, "###########0.00"), "N") & GetStringToPrint(15, "  ", "S") & " " & GetStringToPrint(55, " ", "S") & GetStringToPrint(15, " ", "S") & vbNewLine)
        '                    Print(fnumm, GetStringToPrint(35, "Group Total --> ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(12, Format(grouptaxable, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(groupcgst, "######0.00"), "N") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(12, Format(groupsgst, "######0.00"), "N") & "," & GetStringToPrint(17, Format(groupnet, "######0.00"), "N") & "," & GetStringToPrint(15, "  ", "S") & ", " & GetStringToPrint(55, " ", "S") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
        '                    groupnet = 0
        '                    grouptaxable = 0
        '                    groupcgst = 0
        '                    groupsgst = 0
        '                End If
        '                chkrs1.MoveNext()
        '            End If
        '            If chkrs1.EOF = True Then
        '                Exit Do
        '            End If
        '        Else
        '            Exit Do
        '        End If
        '    Loop
        '    chkrs1.Close()
        Print(fnum, " " & vbNewLine)
        Print(fnumm, " " & vbNewLine)
        Print(fnum, GetStringToPrint(15, "Total -->", "S") & GetStringToPrint(30, Format(ntaxable, "##########0.00"), "N") & GetStringToPrint(30, Format(b2btaxable + b2ctaxable, "##########0.00"), "N") & GetStringToPrint(13, Format(totcgst, "######0.00"), "N") & GetStringToPrint(13, Format(totsgst, "##########0.00"), "N") & GetStringToPrint(15, Format(totnet, "##########0.00"), "N") & vbNewLine)
        Print(fnumm, GetStringToPrint(15, "Total -->", "S") & "," & GetStringToPrint(30, Format(ntaxable, "##########0.00"), "N") & "," & GetStringToPrint(30, Format(b2btaxable + b2ctaxable, "##########0.00"), "N") & "," & GetStringToPrint(13, Format(totcgst, "######0.00"), "N") & "," & GetStringToPrint(13, Format(totsgst, "##########0.00"), "N") & "," & GetStringToPrint(15, Format(totnet, "##########0.00"), "N") & vbNewLine)
        FileClose(fnum)
        FileClose(fnumm)
        Form14.RichTextBox1.LoadFile(Application.StartupPath & "\invoicessummary.dat", RichTextBoxStreamType.PlainText)
        Form14.Show()
        CreatePDF(Application.StartupPath & "\invoicessummary.dat", Application.StartupPath & "\" & TextBox5.Text)
        MsgBox(Application.StartupPath + " \" & TextBox5.Text & ".CSV file is generated")

        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True
        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\" & TextBox5.Text & ".pdf"
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

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        '  ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        '  ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub FrmInvoiceSummary_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
End Class