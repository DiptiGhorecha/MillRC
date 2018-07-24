Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmGSTAdvance
    Dim chkrs As New ADODB.Recordset
    Dim chkrs11 As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim chkrs6 As New ADODB.Recordset
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
    Dim formloaded As Boolean = False
    Dim fnum As Integer
    Dim fnumm As Integer
    Dim pdfpath As String
    Dim strReportFilePath As String
    Private Sub FrmGSTAdvance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        CheckBox1.Checked = False
        'ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'ComboBox4.Text = DateTime.Now.Year
        'ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'ComboBox1.Text = DateTime.Now.Year
        'If DateTime.Now.Month = 1 Then
        '    ComboBox3.Text = DateAndTime.MonthName(12)
        '    ComboBox4.Text = DateTime.Now.Year - 1
        '    '  ComboBox2.Text = DateAndTime.MonthName(12)
        '    '  ComboBox1.Text = DateTime.Now.Year - 1
        'Else
        ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        ComboBox4.Text = DateTime.Now.Year
        '   ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        '  ComboBox1.Text = DateTime.Now.Year
        ' End If
        '        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView1.RowCount > 1 Then
        End If

        TextBox1.Focus()
        formloaded = True
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp

    End Sub
    Public Function showdata(mnth As String, yr As String, mnth1 As String, yr1 As String)
        Try
            Dim DaysInMonth As Integer = 1  'Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            '  MsgBox(DaysInMonth)
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)

            Dim CurrD As DateTime = startP



            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            Dim ST As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"
            da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "RECEIPT")
            DataGridView1.DataSource = ds.Tables("RECEIPT")
            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            DataGridView1.Columns(0).Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = ComboBox3.Text
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = ComboBox4.Text
        End If

        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
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
        'If Trim(ComboBox3.Text) = "" Then
        '    ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'End If
        'If Trim(ComboBox4.Text) = "" Then
        '    ComboBox4.Text = DateTime.Now.Year
        'End If
        'If Trim(ComboBox2.Text) = "" Then
        '    ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'End If
        'If Trim(ComboBox1.Text) = "" Then
        '    ComboBox1.Text = DateTime.Now.Year
        'End If

        'showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        'If DataGridView1.RowCount > 1 Then
        '    TextBox1.Text = DataGridView1.Item(4, 0).Value
        '    TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value
        'Else
        '    TextBox1.Text = ""
        '    TextBox2.Text = ""
        '    TextBox3.Text = ""
        '    TextBox4.Text = ""
        'End If
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = ComboBox3.Text
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = ComboBox4.Text
        End If
        '        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
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
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'If Trim(ComboBox3.Text) = "" Then
        '    ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'End If
        'If Trim(ComboBox4.Text) = "" Then
        '    ComboBox4.Text = DateTime.Now.Year
        'End If
        'If Trim(ComboBox2.Text) = "" Then
        '    ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        'End If
        'If Trim(ComboBox1.Text) = "" Then
        '    ComboBox1.Text = DateTime.Now.Year
        'End If
        ''        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        'showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        'If DataGridView1.RowCount > 1 Then
        '    TextBox1.Text = DataGridView1.Item(4, 0).Value
        '    TextBox2.Text = DataGridView1.Item(4, DataGridView1.RowCount - 1).Value

        'Else
        '    TextBox1.Text = ""
        '    TextBox2.Text = ""
        '    TextBox3.Text = ""
        '    TextBox4.Text = ""
        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DaysInMonth As Integer = 1  'Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '  MsgBox(DaysInMonth)
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        'DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        'Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)

        Dim CurrD As DateTime = startP

        If DataGridView1.RowCount < 1 Then
            MsgBox("No data exist for selected month-Year")
            ComboBox3.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Receipt number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Receipt number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As Int32 = Convert.ToInt32(TextBox1.Text)
        Dim edrec As Int32 = Convert.ToInt32(TextBox2.Text)

        If strrec > edrec Then
            MsgBox("From Receipt number must be less than To Receipt number")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprintgstadvance.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If


        Dim title As String = "GST Advances for " & ComboBox3.Text & " - " & ComboBox4.Text


        If CheckBox2.Checked = True Then   'hsnwise
            chkrs11.Open("Select * FROM GST where HSN_NO='997212' ORDER BY HSN_NO", xcon)
            If chkrs11.BOF = False Then
                chkrs11.MoveFirst()
            End If
            Dim HSN As String
            Dim HSNPRINT As Boolean
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0
            Dim totadvbasic As Double = 0
            Dim totadvgst As Double = 0
            Dim totadvcst As Double = 0
            Do While chkrs11.EOF = False
                HSN = chkrs11.Fields(0).Value
                HSNPRINT = True
                If first Then
                    globalHeader(title, fnum, fnumm)
                    Print(fnum, GetStringToPrint(17, "Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(50, "Party Name", "S") & GetStringToPrint(13, "State Code ", "S") & GetStringToPrint(13, "State", "S") & GetStringToPrint(16, "HSN", "S") & GetStringToPrint(13, "Type", "S") & GetStringToPrint(13, "Tax Rate", "S") & GetStringToPrint(13, "Total Receipt", "N") & GetStringToPrint(15, " Advance Basic ", "N") & GetStringToPrint(15, " Advance CGST ", "N") & GetStringToPrint(15, " Advance SGST ", "N") & GetStringToPrint(15, " Advance IGST ", "N") & GetStringToPrint(15, " Advance Total ", "N") & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(17, "Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(50, "Party Name", "S") & "," & GetStringToPrint(13, "State Code ", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(16, "HSN", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(13, "Tax Rate", "S") & "," & GetStringToPrint(13, "Total Receipt", "N") & "," & GetStringToPrint(15, " Advance Basic ", "N") & "," & GetStringToPrint(15, " Advance CGST ", "N") & "," & GetStringToPrint(15, " Advance SGST ", "N") & "," & GetStringToPrint(15, " Advance IGST ", "N") & "," & GetStringToPrint(15, " Advance Total ", "N") & "," & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)

                    Print(fnum, StrDup(250, "=") & vbNewLine)
                    Print(fnumm, StrDup(250, " ") & vbNewLine)
                    first = False
                    xline = xline + 3
                End If
                HSNPRINT = False
                Dim str As String = "SELECT [RECEIPT].*,[GODOWN].* from [RECEIPT] INNER JOIN [GODOWN] ON ([RECEIPT].[GROUP]=[GODOWN].[GROUP] AND [RECEIPT].[GODWN_NO]=[GODOWN].[GODWN_NO]  where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') And [GODOWN].GST='" & HSN & "' AND [GODOWN].[STATUS]='C' and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO"

                chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO", xcon)
                If chkrs1.BOF = False Then
                    chkrs1.MoveFirst()
                End If
                Dim gtotamt As Double = 0
                Dim gtotadv As Double = 0
                Do While chkrs1.EOF = False
                    If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                        If first Then
                            globalHeader(title, fnum, fnumm)
                            Print(fnum, GetStringToPrint(17, "Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(50, "Party Name", "S") & GetStringToPrint(13, "State Code ", "S") & GetStringToPrint(13, "State", "S") & GetStringToPrint(16, "HSN", "S") & GetStringToPrint(13, "Type", "S") & GetStringToPrint(13, "Tax Rate", "S") & GetStringToPrint(13, "Total Receipt", "N") & GetStringToPrint(15, " Advance Basic ", "N") & GetStringToPrint(15, " Advance CGST ", "N") & GetStringToPrint(15, " Advance SGST ", "N") & GetStringToPrint(15, " Advance IGST ", "N") & GetStringToPrint(15, " Advance Total ", "N") & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(50, "Party Name", "S") & "," & GetStringToPrint(13, "State Code ", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(16, "HSN", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(13, "Tax Rate", "S") & "," & GetStringToPrint(13, "Total Receipt", "N") & "," & GetStringToPrint(15, " Advance Basic ", "N") & "," & GetStringToPrint(15, " Advance CGST ", "N") & "," & GetStringToPrint(15, " Advance SGST ", "N") & "," & GetStringToPrint(15, " Advance IGST ", "N") & "," & GetStringToPrint(15, " Advance Total ", "N") & "," & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)

                            Print(fnum, StrDup(250, "=") & vbNewLine)
                            Print(fnumm, StrDup(250, " ") & vbNewLine)
                            first = False
                            xline = xline + 3
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
                        Dim HSNCD As String = ""
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
                            HSNCD = chkrs4.Fields(37).Value

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
                        Dim last_bldate As DateTime = Nothing
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
                            HSNCD = chkrs2.Fields(11).Value
                            If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                                If FIRSTREC Then
                                    chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs2.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs2.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                    If chkrs6.EOF = False Then
                                        FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                        TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                    Else
                                        FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                        TONO = FROMNO
                                    End If
                                    chkrs6.Close()

                                    FIRSTREC = False
                                Else
                                    TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                End If
                                last_bldate = chkrs2.Fields(4).Value
                                pname = chkrs2.Fields(16).Value
                                pcode1 = chkrs2.Fields(3).Value
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
                        Dim lastbilladjusted As Integer = 0

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
                            If against.Length >= 1 Or lastbilladjusted = 0 Then
                                dtcounter = 1
                            Else
                                dtcounter = 2
                            End If
                            Do Until advanceamt <= 0
                                Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                                If lastbilladjusted = 0 Then
                                    sdt = Convert.ToDateTime(last_bldate)

                                End If

                                If FIRSTREC Then
                                    If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                        FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                        advanceamt = advanceamt - net
                                        TONO = FROMNO
                                        FIRSTREC = False
                                        'last_bldate = chkrs5.Fields(0).Value
                                    End If
                                Else
                                    TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                                    advanceamt = advanceamt - net
                                    dtcounter = dtcounter + 1
                                End If
                            Loop
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If HSNCD.Equals(HSN) Then

                            If CheckBox1.Checked Then
                                If chkrs1.Fields(6).Value = True Then
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(50, pname.Replace(",", " "), "S") & GetStringToPrint(13, " 18 ", "S") & GetStringToPrint(13, " GUJ ", "S") & GetStringToPrint(13, "997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, advanceamtprint, "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(50, pname.Replace(",", " "), "S") & "," & GetStringToPrint(13, " 18 ", "S") & "," & GetStringToPrint(13, " GUJ ", "S") & "," & GetStringToPrint(13, "997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, advanceamtprint, "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    'Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                                    xline = xline + 1

                                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, vbNewLine)
                                    Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                            Else
                                If advanceamtprint > 0 Then
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(50, pname.Replace(",", " "), "S") & GetStringToPrint(13, " 24 ", "S") & GetStringToPrint(13, " GUJ ", "S") & GetStringToPrint(13, "997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, "   18", "S") & GetStringToPrint(15, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(50, pname.Replace(",", " "), "S") & "," & GetStringToPrint(13, " 24 ", "S") & "," & GetStringToPrint(13, " GUJ ", "S") & "," & GetStringToPrint(13, "997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, "   18", "S") & "," & GetStringToPrint(15, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    '
                                    xline = xline + 1
                                    totamt = totamt + chkrs1.Fields(5).Value
                                    totadv = totadv + advanceamtprint
                                    totadvbasic = totadvbasic + Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero)
                                    totadvgst = totadvgst + Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero)
                                    totadvcst = totadvcst + Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero)
                                    gtotamt = gtotamt + chkrs1.Fields(5).Value
                                    gtotadv = gtotadv + advanceamtprint
                                    'Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    'Print(fnumm, vbNewLine)
                                    'Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                            End If
                            End If
                        '    Print(fnum, StrDup(80, "-") & vbNewLine)
                        '  Next
                        If chkrs1.EOF = False Then
                            chkrs1.MoveNext()
                        End If
                        If chkrs1.EOF = True Then
                            Exit Do
                        End If
                    End If
                Loop
                'Print(fnum, StrDup(180, "-") & vbNewLine)
                'Print(fnumm, StrDup(180, "-") & vbNewLine)
                'Print(fnum, GetStringToPrint(17, "Group Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                'Print(fnumm, GetStringToPrint(17, "Group Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                'Print(fnum, StrDup(180, " ") & vbNewLine)
                'Print(fnumm, StrDup(180, " ") & vbNewLine)
                chkrs1.Close()

                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop
            Print(fnum, StrDup(250, " ") & vbNewLine)
            Print(fnum, StrDup(250, "-") & vbNewLine)
            Print(fnumm, StrDup(250, " ") & vbNewLine)
            Print(fnumm, StrDup(250, " ") & vbNewLine)
            ' Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, Format(totamt, "######0.00"), "N") & GetStringToPrint(15, Format(totadvbasic, "######0.00"), "N") & GetStringToPrint(15, Format(totadvgst, "######0.00"), "N") & GetStringToPrint(15, Format(totadvcst, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, Format(totadv, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, Format(totamt, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvbasic, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvgst, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvcst, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs11.Close()
            xcon.Close()

            FileClose(fnum)
            FileClose(fnumm)
            ''''''''''''''''''''''''hsnwise

        End If
        Form17.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprintgstadvance.dat", RichTextBoxStreamType.PlainText)
        Form17.Show()
        CreatePDF(Application.StartupPath & "\Reports\Recprintgstadvance.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        MsgBox(Application.StartupPath + " \Reports\" & TextBox5.Text & ".CSV file is generated")
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"

            Dim pdfPage As PdfPage = pdf.AddPage

            ' pdfPage.Orientation = PdfSharp.PageOrientation.Landscape
            pdfPage.TrimMargins.Left = 15

            pdfPage.Width = 842
            pdfPage.Height = 595

            '  pdf.Pages.RemoveAt(0)
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 7, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1

                line = readFile.ReadLine()

                If line Is Nothing Then
                    Exit While
                Else
                    If counter > 43 Then
                        counter = 1
                        pdfPage = pdf.AddPage()
                        graph = XGraphics.FromPdfPage(pdfPage)
                        font = New XFont("COURIER NEW", 7, XFontStyle.Regular)

                        pdfPage.TrimMargins.Left = 15

                        pdfPage.Width = 842
                        pdfPage.Height = 595
                        yPoint = 60
                    End If
                    'If counter = 1 Or counter = 31 Then
                    '    font = New XFont("COURIER NEW", 14, XFontStyle.Bold)
                    'Else
                    '    font = New XFont("COURIER NEW", 10, XFontStyle.Regular)
                    'End If
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
        Dim DaysInMonth As Integer = 1  'Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '  MsgBox(DaysInMonth)
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        'DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        'Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)

        Dim CurrD As DateTime = startP

        If DataGridView1.RowCount < 1 Then
            MsgBox("No data exist for selected month-Year")
            ComboBox3.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Receipt number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Receipt number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As Int32 = Convert.ToInt32(TextBox1.Text)
        Dim edrec As Int32 = Convert.ToInt32(TextBox2.Text)

        If strrec > edrec Then
            MsgBox("From Receipt number must be less than To Receipt number")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprintgstadvance.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If


        Dim title As String = "GST Advances for " & ComboBox3.Text & " - " & ComboBox4.Text


        If CheckBox2.Checked = True Then   'hsnwise
            chkrs11.Open("Select * FROM GST where HSN_NO='997212' ORDER BY HSN_NO", xcon)
            If chkrs11.BOF = False Then
                chkrs11.MoveFirst()
            End If
            Dim HSN As String
            Dim HSNPRINT As Boolean
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0
            Dim totadvbasic As Double = 0
            Dim totadvgst As Double = 0
            Dim totadvcst As Double = 0
            Do While chkrs11.EOF = False
                HSN = chkrs11.Fields(0).Value
                HSNPRINT = True
                If first Then
                    globalHeader(title, fnum, fnumm)
                    Print(fnum, GetStringToPrint(17, "Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(50, "Party Name", "S") & GetStringToPrint(13, "State Code ", "S") & GetStringToPrint(13, "State", "S") & GetStringToPrint(16, "HSN", "S") & GetStringToPrint(13, "Type", "S") & GetStringToPrint(13, "Tax Rate", "S") & GetStringToPrint(13, "Total Receipt", "N") & GetStringToPrint(15, " Advance Basic ", "N") & GetStringToPrint(15, " Advance CGST ", "N") & GetStringToPrint(15, " Advance SGST ", "N") & GetStringToPrint(15, " Advance IGST ", "N") & GetStringToPrint(15, " Advance Total ", "N") & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(17, "Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(50, "Party Name", "S") & "," & GetStringToPrint(13, "State Code ", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(16, "HSN", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(13, "Tax Rate", "S") & "," & GetStringToPrint(13, "Total Receipt", "N") & "," & GetStringToPrint(15, " Advance Basic ", "N") & "," & GetStringToPrint(15, " Advance CGST ", "N") & "," & GetStringToPrint(15, " Advance SGST ", "N") & "," & GetStringToPrint(15, " Advance IGST ", "N") & "," & GetStringToPrint(15, " Advance Total ", "N") & "," & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)

                    Print(fnum, StrDup(250, "=") & vbNewLine)
                    Print(fnumm, StrDup(250, " ") & vbNewLine)
                    first = False
                    xline = xline + 3
                End If
                HSNPRINT = False
                Dim str As String = "SELECT [RECEIPT].*,[GODOWN].* from [RECEIPT] INNER JOIN [GODOWN] ON ([RECEIPT].[GROUP]=[GODOWN].[GROUP] AND [RECEIPT].[GODWN_NO]=[GODOWN].[GODWN_NO]  where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') And [GODOWN].GST='" & HSN & "' AND [GODOWN].[STATUS]='C' and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO"

                chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO", xcon)
                If chkrs1.BOF = False Then
                    chkrs1.MoveFirst()
                End If
                Dim gtotamt As Double = 0
                Dim gtotadv As Double = 0
                Do While chkrs1.EOF = False
                    If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                        If first Then
                            globalHeader(title, fnum, fnumm)
                            Print(fnum, GetStringToPrint(17, "Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(50, "Party Name", "S") & GetStringToPrint(13, "State Code ", "S") & GetStringToPrint(13, "State", "S") & GetStringToPrint(16, "HSN", "S") & GetStringToPrint(13, "Type", "S") & GetStringToPrint(13, "Tax Rate", "S") & GetStringToPrint(13, "Total Receipt", "N") & GetStringToPrint(15, " Advance Basic ", "N") & GetStringToPrint(15, " Advance CGST ", "N") & GetStringToPrint(15, " Advance SGST ", "N") & GetStringToPrint(15, " Advance IGST ", "N") & GetStringToPrint(15, " Advance Total ", "N") & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(50, "Party Name", "S") & "," & GetStringToPrint(13, "State Code ", "S") & "," & GetStringToPrint(13, "State", "S") & "," & GetStringToPrint(16, "HSN", "S") & "," & GetStringToPrint(13, "Type", "S") & "," & GetStringToPrint(13, "Tax Rate", "S") & "," & GetStringToPrint(13, "Total Receipt", "N") & "," & GetStringToPrint(15, " Advance Basic ", "N") & "," & GetStringToPrint(15, " Advance CGST ", "N") & "," & GetStringToPrint(15, " Advance SGST ", "N") & "," & GetStringToPrint(15, " Advance IGST ", "N") & "," & GetStringToPrint(15, " Advance Total ", "N") & "," & GetStringToPrint(15, " Remarks ", "S") & vbNewLine)

                            Print(fnum, StrDup(250, "=") & vbNewLine)
                            Print(fnumm, StrDup(250, " ") & vbNewLine)
                            first = False
                            xline = xline + 3
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
                        Dim HSNCD As String = ""
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
                            HSNCD = chkrs4.Fields(37).Value

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
                        Dim last_bldate As DateTime = Nothing
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
                            HSNCD = chkrs2.Fields(11).Value
                            If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                                If FIRSTREC Then
                                    chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs2.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs2.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                    If chkrs6.EOF = False Then
                                        FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                        TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                    Else
                                        FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                        TONO = FROMNO
                                    End If
                                    chkrs6.Close()

                                    FIRSTREC = False
                                Else
                                    TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                End If
                                last_bldate = chkrs2.Fields(4).Value
                                pname = chkrs2.Fields(16).Value
                                pcode1 = chkrs2.Fields(3).Value
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
                        Dim lastbilladjusted As Integer = 0

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
                            If against.Length >= 1 Or lastbilladjusted = 0 Then
                                dtcounter = 1
                            Else
                                dtcounter = 2
                            End If
                            Do Until advanceamt <= 0
                                Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                                If lastbilladjusted = 0 Then
                                    sdt = Convert.ToDateTime(last_bldate)

                                End If

                                If FIRSTREC Then
                                    If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                        FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                        advanceamt = advanceamt - net
                                        TONO = FROMNO
                                        FIRSTREC = False
                                        'last_bldate = chkrs5.Fields(0).Value
                                    End If
                                Else
                                    TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                                    advanceamt = advanceamt - net
                                    dtcounter = dtcounter + 1
                                End If
                            Loop
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If HSNCD.Equals(HSN) Then

                            If CheckBox1.Checked Then
                                If chkrs1.Fields(6).Value = True Then
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(13, " 18 ", "S") & GetStringToPrint(13, " GUJ ", "S") & GetStringToPrint(13, "997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, advanceamtprint, "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(13, " 18 ", "S") & "," & GetStringToPrint(13, " GUJ ", "S") & "," & GetStringToPrint(13, "997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, advanceamtprint, "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    'Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                                    xline = xline + 1

                                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, vbNewLine)
                                    Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                            Else
                                If advanceamtprint > 0 Then
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(13, " 24 ", "S") & GetStringToPrint(13, " GUJ ", "S") & GetStringToPrint(13, "997212", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, "   18", "S") & GetStringToPrint(15, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(13, " 24 ", "S") & "," & GetStringToPrint(13, " GUJ ", "S") & "," & GetStringToPrint(13, "997212", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, "   18", "S") & "," & GetStringToPrint(15, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, Format(Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero), "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
                                    '
                                    xline = xline + 1
                                    totamt = totamt + chkrs1.Fields(5).Value
                                    totadv = totadv + advanceamtprint
                                    totadvbasic = totadvbasic + Math.Round((advanceamtprint * 100 / 118), 2, MidpointRounding.AwayFromZero)
                                    totadvgst = totadvgst + Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero)
                                    totadvcst = totadvcst + Math.Round((advanceamtprint - (advanceamtprint * 100 / 118)) / 2, 2, MidpointRounding.AwayFromZero)
                                    gtotamt = gtotamt + chkrs1.Fields(5).Value
                                    gtotadv = gtotadv + advanceamtprint
                                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, vbNewLine)
                                    Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                            End If
                        End If
                        '    Print(fnum, StrDup(80, "-") & vbNewLine)
                        '  Next
                        If chkrs1.EOF = False Then
                            chkrs1.MoveNext()
                        End If
                        If chkrs1.EOF = True Then
                            Exit Do
                        End If
                    End If
                Loop
                'Print(fnum, StrDup(180, "-") & vbNewLine)
                'Print(fnumm, StrDup(180, "-") & vbNewLine)
                'Print(fnum, GetStringToPrint(17, "Group Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                'Print(fnumm, GetStringToPrint(17, "Group Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                'Print(fnum, StrDup(180, " ") & vbNewLine)
                'Print(fnumm, StrDup(180, " ") & vbNewLine)
                chkrs1.Close()

                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop
            Print(fnum, StrDup(250, " ") & vbNewLine)
            Print(fnum, StrDup(250, "-") & vbNewLine)
            Print(fnumm, StrDup(250, " ") & vbNewLine)
            Print(fnumm, StrDup(250, " ") & vbNewLine)
            ' Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(15, Format(totamt, "######0.00"), "N") & GetStringToPrint(15, Format(totadvbasic, "######0.00"), "N") & GetStringToPrint(15, Format(totadvgst, "######0.00"), "N") & GetStringToPrint(15, Format(totadvcst, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(15, Format(totadv, "######0.00"), "N") & GetStringToPrint(15, " ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(15, Format(totamt, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvbasic, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvgst, "######0.00"), "N") & "," & GetStringToPrint(15, Format(totadvcst, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(15, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs11.Close()
            xcon.Close()

            FileClose(fnum)
            FileClose(fnumm)
            ''''''''''''''''''''''''hsnwise

        End If
        Form17.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprintgstadvance.dat", RichTextBoxStreamType.PlainText)
        Form17.Show()
        CreatePDF(Application.StartupPath & "\Reports\Recprintgstadvance.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub FrmGSTAdvance_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            Dim transaction As OleDbTransaction
            transaction = MyConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Dim objcmd As New OleDb.OleDbCommand
            Dim objcmdd As New OleDb.OleDbCommand

            objcmd.Connection = MyConn
            objcmd.Transaction = transaction
            objcmd.CommandType = CommandType.Text
            Dim save As String = "UPDATE [BILL] SET REC_NO=Null, REC_DATE=Null ,ADVANCE=FALSE WHERE INVOICE_NO<>'" & "" & "'"
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()
            objcmdd.Connection = MyConn
            objcmdd.Transaction = transaction
            objcmdd.CommandType = CommandType.Text
            save = "UPDATE [RECEIPT] SET ADJ_AMT=AMOUNT WHERE REC_NO<>0"
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()
            objcmd.CommandText = "delete * from [RECIPTBILL] where REC_NO<>0"
            objcmd.ExecuteNonQuery()
            transaction.Commit()
            MyConn.Close()



            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            Dim AMT As Double = 0
            Dim ADJUSTED As Double = 0
            Dim RCDT As Date
            Dim RCNO As Integer
            Dim REMAINING As Double = 0

            Dim BLDT As Date
            Dim GRP As String
            Dim GDN As String
            Dim PCD As String
            Dim INVNO As String
            Dim hsn As String
            Dim BLAMT As Double
            Dim ADV As Boolean
            Dim gamt As Double
            Dim CGST_AMT As Double
            Dim SGST_AMT As Double
            Dim CGST_RT As Double
            Dim SGST_RT As Double
            Dim FILE_NOtmp As String
            chkrs1.Open("SELECT * FROM BILL ORDER BY BILL_DATE,INVOICE_NO", xcon)

            Do While chkrs1.EOF = False
                BLDT = chkrs1.Fields(4).Value
                GRP = chkrs1.Fields(1).Value
                GDN = chkrs1.Fields(2).Value
                PCD = chkrs1.Fields(3).Value
                INVNO = chkrs1.Fields(0).Value
                BLAMT = chkrs1.Fields(10).Value
                hsn = chkrs1.Fields(11).Value
                gamt = chkrs1.Fields(5).Value
                CGST_AMT = chkrs1.Fields(7).Value
                SGST_AMT = chkrs1.Fields(9).Value
                CGST_RT = chkrs1.Fields(6).Value
                SGST_RT = chkrs1.Fields(8).Value
                FILE_NOtmp = chkrs1.Fields(12).Value
                Dim str As String = "SELECT * FROM RECEIPT WHERE GROUP='" & GRP & "' AND GODWN_NO='" & GDN & "' AND ADJ_AMT>0 ORDER BY [GROUP],GODWN_NO,REC_DATE,REC_NO"
                chkrs2.Open("SELECT TOP 1 * FROM RECEIPT WHERE [GROUP]='" & GRP & "' AND GODWN_NO='" & GDN & "' AND ADJ_AMT>0 ORDER BY [GROUP],GODWN_NO,REC_DATE,REC_NO", xcon)
                If chkrs2.EOF = False Then
                    ' chkrs2.MoveFirst()
                    RCDT = chkrs2.Fields(3).Value
                    RCNO = chkrs2.Fields(4).Value
                    REMAINING = chkrs2.Fields(13).Value - BLAMT
                    If (RCDT < BLDT) Then
                        ADV = True
                    Else
                        ADV = False
                    End If
                    MyConn = New OleDbConnection(connString)
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    transaction = MyConn.BeginTransaction(IsolationLevel.ReadUncommitted)
                    Dim objcmd1 As New OleDb.OleDbCommand
                    objcmd1.Connection = MyConn
                    objcmd1.Transaction = transaction
                    objcmd1.CommandType = CommandType.Text
                    If ADV = False Then
                        objcmd1.CommandText = "INSERT INTO [RECIPTBILL](REC_NO,INVOICE_NO,AMT,REC_DATE) VALUES('" & RCNO & "','" & INVNO & "','" & BLAMT & "',Format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy'))"
                        objcmd1.ExecuteNonQuery()
                    End If
                    save = "UPDATE [BILL] SET REC_NO=" & RCNO & ", REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy'),ADVANCE=" & ADV & " WHERE INVOICE_NO='" & INVNO & "' AND [GROUP]='" & GRP & "' AND GODWN_NO='" & GDN & "'"  ''' sorry about that
                    objcmd1.CommandText = save
                    objcmd1.ExecuteNonQuery()
                    ' str = "UPDATE [RECEIPT] SET ADJ_AMT=" & REMAINING & " WHERE REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy') AND REC_NO=" & RCNO
                    objcmd1.CommandText = "UPDATE [RECEIPT] SET ADJ_AMT=" & REMAINING & " WHERE REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy') AND REC_NO=" & RCNO  ' sorry about that
                    objcmd1.ExecuteNonQuery()
                    transaction.Commit()
                    objcmd1.Dispose()
                    MyConn.Close()
                    '''''' System.Threading.Thread.Sleep(500)
                End If
                'Dim i As Integer
                'For i = 1 To 5000

                'Next
                chkrs2.Close()
                ''''''''''''''''''''''''''''''''invoice reprint start
                fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month))
                End If
                FileOpen(fnum, Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month) & "\" & INVNO & ".dat", OpenMode.Output)
                Dim TENANT_CODE As String
                Dim TENANT_NAME As String
                Dim T_ADDREASS As String
                Dim TAD1 As String
                Dim TAD2 As String
                Dim TAD3 As String
                Dim TCITY As String
                Dim TSTATE As String
                Dim STATE_CODE As String
                Dim TEMAIL As String
                Dim TGST As String
                Dim CGST_TAXAMT As Double
                Dim SGST_TAXAMT As Double
                Dim CGST_RATE As Double
                Dim SGST_RATE As Double
                Dim gst As Double
                Dim gst_amt As Double
                TENANT_CODE = PCD

                chkrs11.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
                If chkrs11.EOF = False Then
                    TENANT_NAME = LTrim(chkrs11.Fields(1).Value)

                    If IsDBNull(chkrs11.Fields(2).Value) Then
                        TAD1 = ""
                    Else
                        If Trim(chkrs11.Fields(2).Value).Equals("") Then
                            TAD1 = ""
                        Else
                            TAD1 = chkrs11.Fields(2).Value.replace("& vbLf & vbLf", "")
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(3).Value) Then
                        TAD2 = ""
                    Else
                        If Trim(chkrs11.Fields(3).Value).Equals("") Then
                            TAD2 = ""
                        Else
                            TAD2 = chkrs11.Fields(3).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(4).Value) Then
                        TAD3 = ""
                    Else
                        If Trim(chkrs11.Fields(4).Value).Equals("") Then
                            TAD3 = ""
                        Else
                            TAD3 = chkrs11.Fields(4).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(5).Value) Then
                        TCITY = ""
                    Else
                        If Trim(chkrs11.Fields(5).Value).Equals("") Then
                            TCITY = ""
                        Else
                            TCITY = chkrs11.Fields(5).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(17).Value) Then
                        TSTATE = ""
                    Else
                        If Trim(chkrs11.Fields(17).Value).Equals("") Then
                            TSTATE = ""
                        Else
                            TSTATE = chkrs11.Fields(17).Value
                        End If
                    End If
                    STATE_CODE = "24"

                    If IsDBNull(chkrs11.Fields(18).Value) Then
                        TEMAIL = ""
                    Else
                        TEMAIL = chkrs11.Fields(18).Value
                    End If
                    If IsDBNull(chkrs11.Fields(19).Value) Then
                        TGST = ""
                    Else
                        TGST = chkrs11.Fields(19).Value
                    End If
                End If
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, "BILLED TO :" & Space(35) & "[  ] Original for Recepient" & vbNewLine)
                Print(fnum, "           " & Space(35) & "[  ] Duplicate for Supplier" & vbNewLine)
                Print(fnum, GetStringToPrint(45, TENANT_NAME, "S") & " " & StrDup(27, "-") & vbNewLine)
                If TAD1 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD1, "S") & vbNewLine)
                End If
                If TAD2 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD2, "S") & vbNewLine)
                End If
                If TAD3 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD3, "S") & vbNewLine)
                End If
                If TCITY <> "" Then
                    Print(fnum, GetStringToPrint(50, TCITY, "S") & vbNewLine)
                End If
                '  If TSTATE <> "" Then
                Print(fnum, GetStringToPrint(30, "STATE :" & TSTATE, "S") & Space(15) & "GODOWN NO.  :" & GetStringToPrint(20, GRP & GDN, "S") & vbNewLine)
                '  End If
                ' If TGST <> "" Then
                Print(fnum, "GSTIN :" & GetStringToPrint(30, TGST, "S") & Space(8) & "INVOICE NO. :" & GetStringToPrint(35, "GO-" & INVNO, "S") & vbNewLine)
                '  End If
                '  If TEMAIL <> "" Then
                Print(fnum, "EMAIL ID:" & GetStringToPrint(33, TEMAIL, "S") & Space(3) & "INVOICE DATE:" & GetStringToPrint(20, BLDT.ToString("dd/MM/yyyy"), "S") & vbNewLine)
                '   End If
                chkrs11.Close()

                Print(fnum, StrDup(30, " ") & vbNewLine)
                Print(fnum, StrDup(30, " ") & "TAX INVOICE FOR SERVICES" & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(8, "HSN", "S") & GetStringToPrint(28, "HSN DESCRIPTION", "S") & GetStringToPrint(30, "DESCRIPTION OF SERVICES", "S") & GetStringToPrint(19, "AMOUNT", "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                If IsDBNull(hsn) Then
                    Print(fnum, GetStringToPrint(7, "", "S"))
                Else
                    Print(fnum, GetStringToPrint(7, hsn, "S"))

                End If
                Print(fnum, GetStringToPrint(28, " Rental Or Leasing Services ", "S"))
                Print(fnum, GetStringToPrint(41, " Rent for property from " & "1st " & MonthName(BLDT.Month) & "," & BLDT.Year.ToString, "S"))
                '  ComboBox1.Text = DateTimePicker1.Value.Month
                '   ComboBox2.Text = DateTimePicker1.Value.Year

                ' AMT = TextBox4.Text
                Print(fnum, GetStringToPrint(9, Format(gamt, "#####0.00"), "N") & vbNewLine)
                Dim ENDDAY As String
                ENDDAY = DateTime.DaysInMonth(BLDT.Year, BLDT.Month).ToString
                If ENDDAY = "31" Then
                    ENDDAY = "31st"
                Else
                    ENDDAY = ENDDAY + "th"
                End If
                Print(fnum, GetStringToPrint(7, "", "S"))
                Print(fnum, GetStringToPrint(28, " Involving Own Or Leased ", "S"))
                Print(fnum, GetStringToPrint(35, " to " & ENDDAY & " " & MonthName(BLDT.Month) & "," & BLDT.Year.ToString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S"))
                If IsDBNull(hsn) Or hsn.Equals("997211") Then
                    Print(fnum, GetStringToPrint(29, " Residential Property ", "S"))
                    '    Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                Else
                    Print(fnum, GetStringToPrint(29, " Non-residential Property ", "S"))
                    '   Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                End If

                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "TAXABLE AMOUNT :", "S") & GetStringToPrint(10, Format(gamt, "#####0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)

                'CGST_TAXAMT = amt * CGST_RATE / 100
                'CGST_TAXAMT = Math.Round(CGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                'SGST_TAXAMT = Math.Round(SGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
                Dim net As Double
                Dim rnd As Integer


                net = BLAMT
                CGST_TAXAMT = CGST_AMT
                SGST_TAXAMT = SGST_AMT
                CGST_RATE = CGST_RT
                SGST_RATE = SGST_RT

                Print(fnum, Space(58) & GetStringToPrint(17, "CGST@ " & CGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(CGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "SGST@ " & SGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(SGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(58, " ") & StrDup(27, "-") & vbNewLine)

                Print(fnum, Space(58) & GetStringToPrint(17, "NET AMOUNT     :", "S") & GetStringToPrint(10, Format(net, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(net)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(50, inwordd.Length - 50)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(23) & GetStringToPrint(51, inword1, "S") & vbNewLine)
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If

                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(40) & GetStringToPrint(45, "For Motilal Hirabhai Estate & Warehouse Ltd.", "S") & vbNewLine)
                If ADV = True Then
                    Print(fnum, GetStringToPrint(23, " ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(24, "Received as advance on:", "S") & GetStringToPrint(19, RCDT, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Receipt No.:", "S") & GetStringToPrint(19, "GST-" + Convert.ToString(RCNO), "S") & vbNewLine)
                Else
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                End If
                Print(fnum, Space(40) & GetStringToPrint(45, "Authorised Signatory", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Subject to Ahmedabad jurisdiction.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "This is computer generated invoice.", "S") & vbNewLine)

                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Ref. No. :" + FILE_NOtmp, "S") & vbNewLine)
                FileClose(fnum)
                pdfpath = Application.StartupPath & "\Invoices\pdf\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month)
                strReportFilePath = Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month) & "\" & INVNO & ".dat"
                CreatePDFNEW(strReportFilePath, INVNO)
                ''''''''''''''''''''''''''''''''invoice reprint end
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
            Loop
            chkrs1.Close()
            xcon.Close()

        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub
    Private Function CreatePDFNEW(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New System.IO.StreamReader(strReportFilePath)
            Dim yPoint As Integer = 0

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 15
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)


            Dim font As XFont = New XFont("COURIER NEW", 9, XFontStyle.Regular)


            While True
                line = readFile.ReadLine()
                If line Is Nothing Then
                    Exit While
                Else
                    graph.DrawString(line, font, XBrushes.Black,
                    New XRect(50, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = pdfpath & "\" & invoice_no & ".pdf"

            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

End Class