﻿Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
''' <summary>
''' tables used - receipt,reciptbill,godown,bill,party,rent,gst
''' this is form to accept inputs from user to view/print receipt CHECKLIST
''' Form17.vb is used to hold report view
''' </summary>
Public Class FrmRecChecklist
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

    Private Sub FrmRecChecklist_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '''''''set position of the form
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        CheckBox1.Checked = False   '''''''by default keep advance receipt only checkbox is unchecked
        '''''set from month-year and to month-year combobox selection one month less than current month
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
        ''''show data in receipt datagrid for selected criteria using receipt table
        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
        If DataGridView1.RowCount > 1 Then
        End If

        TextBox1.Focus()
        formloaded = True
    End Sub
    Public Function showdata(mnth As String, yr As String, mnth1 As String, yr1 As String)
        ''''show data in receipt datagrid for selected criteria using receipt table
        Try
            Dim DaysInMonth As Integer = 1
            Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
            DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
            Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
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
        ''''show data in receipt datagrid for selected criteria using receipt table when user select month for From Month combobox
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = DateTime.Now.Year
        End If

        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
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
        ''''show data in receipt datagrid for selected criteria using receipt table when user select month for To Month combobox
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = DateTime.Now.Year
        End If

        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
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
        ''''show data in receipt datagrid for selected criteria using receipt table when user select month for From Month combobox
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = DateTime.Now.Year
        End If
        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
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
        ''''show data in receipt datagrid for selected criteria using receipt table when user select month for To Month combobox
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        If Trim(ComboBox2.Text) = "" Then
            ComboBox2.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox1.Text) = "" Then
            ComboBox1.Text = DateTime.Now.Year
        End If
        showdata(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox1.Text)
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()    ''''close the form
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        ''''allow only numeric for receipt number
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        ''''allow only numeric for receipt number
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '''''report view
        Dim DaysInMonth As Integer = 1  'Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '''''''start date
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '''''''''end date
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
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

        Dim strrec As Int32 = Convert.ToInt32(TextBox1.Text)     '''''from receipt number
        Dim edrec As Int32 = Convert.ToInt32(TextBox2.Text)      '''''to receipt nmber

        If strrec > edrec Then
            MsgBox("From Receipt number must be less than To Receipt number")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 '''''''''Get FreeFile No.'''''''''''this is for generating comma separated file 
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprintchecklist.dat", OpenMode.Output)
        FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If


        Dim title As String = "Receipt Checklist"
        If CheckBox2.Checked Then
            title = title & " HSN Wise "
        End If
        If CheckBox1.Checked Then
            title = title & " - Advance Only"
        End If

        If CheckBox2.Checked = True Then   '''''''''hsnwise report start
            chkrs11.Open("Select * FROM GST ORDER BY HSN_NO", xcon)
            If chkrs11.BOF = False Then
                chkrs11.MoveFirst()
            End If
            Dim HSN As String
            Dim HSNPRINT As Boolean
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0

            ''''loop through GST type
            Do While chkrs11.EOF = False
                HSN = chkrs11.Fields(0).Value
                HSNPRINT = True
                If first Then
                    globalHeader(title, fnum, fnumm)

                    ''''''''report header
                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & "," & GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & "," & GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                    Print(fnum, StrDup(180, "=") & vbNewLine)
                    Print(fnumm, StrDup(180, "=") & vbNewLine)
                    first = False
                    xline = xline + 3
                End If
                Print(fnum, GetStringToPrint(35, "HSN Number : " & chkrs11.Fields(0).Value, "S") & GetStringToPrint(75, chkrs11.Fields(1).Value, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(35, "HSN Number : " & chkrs11.Fields(0).Value, "S") & GetStringToPrint(75, chkrs11.Fields(1).Value, "S") & vbNewLine)
                HSNPRINT = False
                xline = xline + 1

                ''''this string is for debug purpose only
                Dim str As String = "SELECT [RECEIPT].*,[GODOWN].* from [RECEIPT] INNER JOIN [GODOWN] ON ([RECEIPT].[GROUP]=[GODOWN].[GROUP] AND [RECEIPT].[GODWN_NO]=[GODOWN].[GODWN_NO]  where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') And [GODOWN].GST='" & HSN & "' AND [GODOWN].[STATUS]='C' and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO"

                chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO", xcon)
                If chkrs1.BOF = False Then
                    chkrs1.MoveFirst()
                End If
                Dim gtotamt As Double = 0
                Dim gtotadv As Double = 0
                Do While chkrs1.EOF = False
                    If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                        If first Then ''''''''report header on new [age
                            Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & "," & GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & "," & GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(73, "Adjusted Bill No.", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, "Adjusted Bill No.", "S") & vbNewLine)
                            Print(fnum, StrDup(180, "=") & vbNewLine)
                            Print(fnumm, StrDup(180, "") & vbNewLine)
                            first = False
                            xline = xline + 3
                        End If

                        '''''''''''''''''''''''''''''''''godown details'''''''''''''''''''''''''''''''''''
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

                            '''''''''take rent from rent table for group+godwn_no+p_code
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
                            ''''''take gst rates for gst type from gst table
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
                            CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                            SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                        End If
                        chkrs4.Close()

                        '''''''''''''''''''''''''''''''''godown details'''''''''''''''''''''''''''''''''''

                        '''''''''''''''''''against bill and period start''''''''''''''
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

                        Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                        chkrs2.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                        Do While chkrs2.EOF = False
                            HSNCD = chkrs2.Fields(11).Value
                            If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                                '''''''''''''''''''''''''''''get from month and to month for adjustment
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
                                '''''''''''''''''''''''''''''get from month and to month for adjustment
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
                                    End If
                                Else
                                    TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                                    advanceamt = advanceamt - net
                                    dtcounter = dtcounter + 1
                                End If
                            Loop
                        End If

                        ''''''''''''''find out if any advance is left after adjustment end

                        If HSNCD.Equals(HSN) Then  '''''if HSN number equals current hsn from GST table, do printing
                            totamt = totamt + chkrs1.Fields(5).Value
                            totadv = totadv + advanceamtprint
                            gtotamt = gtotamt + chkrs1.Fields(5).Value
                            gtotadv = gtotadv + advanceamtprint
                            If CheckBox1.Checked Then   ''''check for advance receipt only
                                If chkrs1.Fields(6).Value = True Then    '''''only print receipt details having advance
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                                    xline = xline + 1
                                    Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & "," & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & ",")
                                    xline = xline + 1

                                    Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(73, against, "S") & ",")
                                    xline = xline + 1
                                    If Trim(against1).Equals("") Then
                                    Else
                                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against1, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(73, against1, "S") & ",")
                                        xline = xline + 1
                                    End If
                                    If Trim(against2).Equals("") Then
                                    Else
                                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against2, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(73, against2, "S") & ",")
                                        xline = xline + 1
                                    End If
                                    If Trim(against3).Equals("") Then
                                    Else
                                        Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against3, "S") & vbNewLine)
                                        Print(fnumm, GetStringToPrint(73, against3, "S") & ",")
                                        xline = xline + 1
                                    End If
                                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    Print(fnumm, vbNewLine)
                                    Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                            Else   '''''print all receipts including advance receipts
                                Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & "," & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & ",")
                                xline = xline + 1

                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(73, against, "S") & ",")
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(73, against1, "S") & ",")
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(73, against2, "S") & ",")
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(73, against3, "S") & ",")
                                    xline = xline + 1
                                End If
                                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                Print(fnumm, vbNewLine)
                                Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                xline = xline + 1
                            End If
                        End If
                        If chkrs1.EOF = False Then
                            chkrs1.MoveNext()
                        End If
                        If chkrs1.EOF = True Then
                            Exit Do
                        End If
                    End If
                Loop    '''''''''receipt table
                Print(fnum, StrDup(180, "-") & vbNewLine)
                Print(fnumm, StrDup(180, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(17, "Group Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(17, "Group Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                Print(fnum, StrDup(180, " ") & vbNewLine)
                Print(fnumm, StrDup(180, " ") & vbNewLine)
                chkrs1.Close()




                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop    '''''''''gst table 
            Print(fnum, StrDup(180, " ") & vbNewLine)
            Print(fnum, StrDup(180, "-") & vbNewLine)
            Print(fnumm, StrDup(180, " ") & vbNewLine)
            Print(fnumm, StrDup(180, "-") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs11.Close()
            xcon.Close()

            FileClose(fnum)
            FileClose(fnumm)
            ''''''''''''''''''''''''hsnwise report end
        Else
            Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between #" & startP & "# AND #" & endP & "#  order by YEAR([RECEIPT].REC_DATE)+[RECEIPT].REC_NO"

            chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY')  and REC_NO>=" & strrec & " AND rec_no<=" & edrec & " order by [RECEIPT].REC_DATE,[RECEIPT].REC_NO", xcon)
            If chkrs1.BOF = False Then
                chkrs1.MoveFirst()
            End If
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0
            Do While chkrs1.EOF = False    ''''receipt loop start
                If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                    If first Then
                        globalHeader(title, fnum, fnumm)
                        Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & ",")
                        Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & ",")
                        Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, "Adjusted Bill No.", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(73, "Adjusted Bill No.", "S") & vbNewLine)
                        Print(fnum, StrDup(180, "=") & vbNewLine)
                        Print(fnumm, StrDup(180, "=") & vbNewLine)
                        first = False
                        xline = xline + 3
                    End If

                    '''''''''''''''''''godown detail start''''''''''''
                    chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
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
                        CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                        SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    End If
                    chkrs4.Close()
                    '''''''''''''''''''godown detail end''''''''''''

                    '''''''''''''''''''against bill and period start'''''''''''
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
                    Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs2.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                    Do While chkrs2.EOF = False
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
                                TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                            End If
                            last_bldate = chkrs2.Fields(4).Value
                            pname = chkrs2.Fields(16).Value
                            pcode1 = chkrs2.Fields(3).Value  ''LAST CHANGE DONE BY DIPTI
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
                                End If
                            Else
                                TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                                advanceamt = advanceamt - net
                                dtcounter = dtcounter + 1
                            End If
                        Loop
                    End If
                    ''''''''''''''find out if any advance is left after adjustment end

                    totamt = totamt + chkrs1.Fields(5).Value
                    totadv = totadv + advanceamtprint
                    If CheckBox1.Checked Then     ''''''check for advance receipt only
                        If chkrs1.Fields(6).Value = True Then   ''''print only receipt will advance
                            Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                            xline = xline + 1
                            Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & "," & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & ",")
                            xline = xline + 1

                            Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(63, against, "S") & ",")
                            xline = xline + 1
                            If Trim(against1).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against1, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(73, against1, "S") & ",")
                                xline = xline + 1
                            End If
                            If Trim(against2).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against2, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(73, against2, "S") & ",")
                                xline = xline + 1
                            End If
                            If Trim(against3).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against3, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(73, against3, "S") & ",")
                                xline = xline + 1
                            End If
                            Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                            Print(fnumm, vbNewLine)
                            Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                            xline = xline + 1
                        End If
                    Else       ''''print all receipts
                        Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & ",")

                        xline = xline + 1
                        Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & "," & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & ",")
                        xline = xline + 1

                        Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against, "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(73, against, "S") & ",")
                        xline = xline + 1
                        If Trim(against1).Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against1, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(73, against1, "S") & ",")
                            xline = xline + 1
                        End If
                        If Trim(against2).Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against2, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(73, against2, "S") & ",")
                            xline = xline + 1
                        End If
                        If Trim(against3).Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(73, against3, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(73, against3, "S") & ",")
                            xline = xline + 1
                        End If
                        Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(1, " ", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                        xline = xline + 1
                    End If
                    If chkrs1.EOF = False Then
                        chkrs1.MoveNext()
                    End If
                    If chkrs1.EOF = True Then
                        Exit Do
                    End If
                End If
            Loop    '''''receipt loop end
            Print(fnum, StrDup(180, "-") & vbNewLine)
            Print(fnumm, StrDup(180, "-") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs1.Close()
            xcon.Close()
            FileClose(fnum)
            FileClose(fnumm)
        End If

        ''''display created .dat file in report view form
        Form17.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprintchecklist.dat", RichTextBoxStreamType.PlainText)

        Form17.Show()    '''''show report
        CreatePDF(Application.StartupPath & "\Reports\Recprintchecklist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)   ''''generate pdf file from .dat file
        MsgBox(Application.StartupPath + " \Reports\" & TextBox5.Text & ".CSV file is generated")
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        '''''create pdf file from .dat file
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
        '''''''''same logic as button1_click event. Only one difference in this module we are not generating .csv file
        '''''''''at the end of method printing a pdf file
        '''''''''print report
        Dim DaysInMonth As Integer = 1
        '''''start date
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox1.Text, DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        ''''end date
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox1.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
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
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\\Recprintchecklist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim title As String = "Receipt Checklist"
        If CheckBox2.Checked Then
            title = title & " HSN Wise "
        End If
        If CheckBox1.Checked Then
            title = title & " - Advance Only"
        End If
        If CheckBox2.Checked = True Then   'hsnwise
            chkrs11.Open("SELECT * FROM GST ORDER BY HSN_NO", xcon)
            If chkrs11.BOF = False Then
                chkrs11.MoveFirst()
            End If
            Dim HSN As String
            Dim HSNPRINT As Boolean
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0
            Do While chkrs11.EOF = False
                HSN = chkrs11.Fields(0).Value
                HSNPRINT = True
                If first Then
                    globalHeader(title, fnum, 0)
                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                    Print(fnum, StrDup(180, "=") & vbNewLine)
                    first = False
                    xline = xline + 3
                End If
                Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs11.Fields(0).Value, "S") & GetStringToPrint(75, chkrs11.Fields(1).Value, "S") & vbNewLine)
                HSNPRINT = False
                xline = xline + 1


                Dim str As String = "SELECT [RECEIPT].*,[GODOWN].* from [RECEIPT] INNER JOIN [GODOWN] ON ([RECEIPT].[GROUP]=[GODOWN].[GROUP] AND [RECEIPT].[GODWN_NO]=[GODOWN].[GODWN_NO]  where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') And [GODOWN].GST='" & HSN & "' AND [GODOWN].[STATUS]='C' and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"

                chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') and [RECEIPT].REC_NO>=" & strrec & " AND [RECEIPT].rec_no<=" & edrec & " order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
                If chkrs1.BOF = False Then
                    chkrs1.MoveFirst()
                End If
                Dim gtotamt As Double = 0
                Dim gtotadv As Double = 0
                Do While chkrs1.EOF = False
                    If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                        If first Then

                            Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                            Print(fnum, StrDup(180, "=") & vbNewLine)
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
                                Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                                If FIRSTREC Then
                                    If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                        FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                        advanceamt = advanceamt - net
                                        TONO = FROMNO
                                        FIRSTREC = False
                                        'last_bldate = chkrs5.Fields(0).Value
                                    End If
                                Else
                                    TONO = MonthName(sdt.AddMonths(dtcounter).Month, False) & "-" & sdt.AddMonths(dtcounter).Year
                                    advanceamt = advanceamt - net
                                    dtcounter = dtcounter + 1
                                End If
                            Loop
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If HSNCD.Equals(HSN) Then
                            totamt = totamt + chkrs1.Fields(5).Value
                            totadv = totadv + advanceamtprint
                            gtotamt = gtotamt + chkrs1.Fields(5).Value
                            gtotadv = gtotadv + advanceamtprint
                            If CheckBox1.Checked Then
                                If chkrs1.Fields(6).Value = True Then
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
                Print(fnum, StrDup(180, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(17, "Group Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                Print(fnum, StrDup(180, " ") & vbNewLine)
                chkrs1.Close()




                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop
            Print(fnum, StrDup(180, " ") & vbNewLine)
            Print(fnum, StrDup(180, "-") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs11.Close()
            xcon.Close()

            FileClose(fnum)
            ''''''''''''''''''''''''hsnwise
        Else
            Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between #" & startP & "# AND #" & endP & "#  order by YEAR([RECEIPT].REC_DATE)+[RECEIPT].REC_NO"

            chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY')  and REC_NO>=" & strrec & " AND rec_no<=" & edrec & " order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
            If chkrs1.BOF = False Then
                chkrs1.MoveFirst()
            End If
            Dim first As Boolean = True
            Dim totamt As Double = 0
            Dim totadv As Double = 0
            Do While chkrs1.EOF = False
                If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                    If first Then
                        globalHeader(title, fnum, 0)
                        Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                        Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                        Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                        Print(fnum, StrDup(180, "=") & vbNewLine)
                        first = False
                        xline = xline + 3
                    End If

                    '''''''''''''''''''''''godown detail start'''''''''''''''''''''''''''''''''''''''''''''
                    chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
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
                        CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                        SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    End If
                    chkrs4.Close()
                    '''''''''''''''''''''''godown detail end'''''''''''''''''''''''''''''''''''''''''''''

                    '''''''''''''''''''against bill and period start''''''''''''
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

                    Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs2.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                    Do While chkrs2.EOF = False

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
                                last_bldate = chkrs5.Fields(0).Value
                            End If
                            chkrs5.Close()
                        End If
                        Dim dtcounter As Integer = 1
                        Do Until advanceamt <= 0
                            Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                            If FIRSTREC Then
                                If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                    FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                    advanceamt = advanceamt - net
                                    TONO = FROMNO
                                    FIRSTREC = False
                                End If
                            Else
                                TONO = MonthName(sdt.AddMonths(dtcounter).Month, False) & "-" & sdt.AddMonths(dtcounter).Year
                                advanceamt = advanceamt - net
                                dtcounter = dtcounter + 1
                            End If
                        Loop
                    End If
                    ''''''''''''''find out if any advance is left after adjustment end

                    totamt = totamt + chkrs1.Fields(5).Value
                    totadv = totadv + advanceamtprint
                    If CheckBox1.Checked Then
                        If chkrs1.Fields(6).Value = True Then
                            Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)

                            xline = xline + 1
                            Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                            xline = xline + 1

                            Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                            xline = xline + 1
                            If Trim(against1).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                xline = xline + 1
                            End If
                            If Trim(against2).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                xline = xline + 1
                            End If
                            If Trim(against3).Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                xline = xline + 1
                            End If
                            Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
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
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                            xline = xline + 1
                        End If
                        If Trim(against2).Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                            xline = xline + 1
                        End If
                        If Trim(against3).Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                            xline = xline + 1
                        End If
                        Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                        xline = xline + 1
                    End If
                    If chkrs1.EOF = False Then
                        chkrs1.MoveNext()
                    End If
                    If chkrs1.EOF = True Then
                        Exit Do
                    End If
                End If
            Loop
            Print(fnum, StrDup(180, "-") & vbNewLine)
            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            chkrs1.Close()
            xcon.Close()
            FileClose(fnum)
        End If

        Me.PrintDialog1.PrintToFile = False
        If Me.PrintDialog1.ShowDialog() = DialogResult.OK Then
        End If
        Form17.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprintchecklist.dat", RichTextBoxStreamType.PlainText)
        CreatePDF(Application.StartupPath & "\Reports\Recprintchecklist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Form17.Show()
        '''''''''same logic as button1_click event. Only one difference in this module we are not generating .csv file

        ''''send pdf file to default printer
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub FrmRecChecklist_Move(sender As Object, e As EventArgs) Handles Me.Move
        ''''keep form position fix
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
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
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

End Class