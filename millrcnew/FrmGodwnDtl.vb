﻿Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
''' <summary>
''' tables used - godown,party,bill,gst,rent,group,gdtrans,clgdwn
''' this is form to accept inputs from user to view/print godown detail
''' FrmGdnDtlView.vb is used to hold report view 
''' </summary>
Public Class FrmGodwnDtl
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs222 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim chkrs6 As New ADODB.Recordset
    Dim chkrs7 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"

    Dim MyConn As OleDbConnection
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim strReportFilePath As String
    Dim GrpAddCorrect As String
    Dim blnTranStart As Boolean
    Dim oldName As String
    Dim ok As Boolean
    Private bValidatepcode As Boolean = True
    Private bValidatepname As Boolean = True
    Private BVALIDATEEMAIL As Boolean = True
    Dim formloaded As Boolean = False
    Private indexorder As String = "[GODOWN].P_NAME"
    Dim fnum As Integer
    Dim fnumm As Integer
    Dim chkrs11 As New ADODB.Recordset
    Dim chkrs22 As New ADODB.Recordset
    Dim chkrs33 As New ADODB.Recordset
    Dim chkrs44 As New ADODB.Recordset
    Dim chkrs55 As New ADODB.Recordset
    Dim chkrs66 As New ADODB.Recordset
    Dim chkrs77 As New ADODB.Recordset
    Private Sub FrmGodwnDtl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''''''set position of the form
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        For Each column As DataGridViewColumn In DataGridView2.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        fillgroupcombo()       ''''''fill godown group combobox with group using group table
        ShowData(ComboBox1.Text, TextBox1.Text)     '''''show data in godown datagrid for selected group
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(3, 0).Value          '''''''assign value of godwn_no from 1st row of datagrid to godown textbox
            ComboBox1.Text = GetValue(DataGridView2.Item(0, 0).Value)     '''''''assign value of group from 1st row of datagrid to group combobox
            ComboBox1.SelectedIndex = ComboBox1.FindStringExact(ComboBox1.Text)
        End If
        Dim iDate As String
        Dim fDate As DateTime
        Dim oDate As String
        Dim foDate As DateTime
        If (Month(Convert.ToDateTime(Date.Now)) >= 4) Then
            iDate = "01/04/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))     ''''''finnce year start date
            fDate = Convert.ToDateTime(iDate)
            oDate = "31/03/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) + 1))    '''''''''''''finnce year end date
            foDate = Convert.ToDateTime(oDate)
            DateTimePicker3.Value = foDate          ''''' to invoice number
            DateTimePicker2.Value = fDate           '''''from invoice number
            DateTimePicker4.Value = foDate          ''''to receipt number
            DateTimePicker5.Value = fDate            '''''' from receipt number
        Else
            iDate = "01/04/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) - 1))     ''''''finnce year start date
            fDate = Convert.ToDateTime(iDate)
            oDate = "31/03/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))     '''''''''''''finnce year end date
            foDate = Convert.ToDateTime(oDate)
            DateTimePicker3.Value = foDate
            DateTimePicker2.Value = fDate
            DateTimePicker4.Value = foDate
            DateTimePicker5.Value = fDate
        End If
        ComboBox1.Focus()
        formloaded = True
    End Sub
    Private Sub ShowData(grp As String, frGdn As String)
        ''''''''show godown details in data grid view uing godown,party table
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GODOWN].[STATUS]='C' order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Function fillgroupcombo()
        '''''fill godown group combo box with g_code using group table
        Try
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
        Catch ex As Exception
            MessageBox.Show("Group combo fill :" & ex.Message)
        End Try
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '''''''report view
        If (TextBox1.Text = "") Then
            MsgBox("Please enter Tenant code .......")
            TextBox1.Focus()
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 '''''''''Get FreeFile No. to write .csv file '''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        ''''''get data for selected group and godown number from godown table
        chkrs222.Open("SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' AND [GODOWN].[STATUS]='C' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D", xcon)
        If chkrs222.BOF = False Then
            chkrs222.MoveFirst()
        End If
        Dim pcdd As String = ""
        Do While chkrs222.EOF = False
            pcdd = chkrs222.Fields(1).Value
            Exit Do
        Loop
        chkrs222.Close()
        ''''''get data for tenant code of godown table from party table
        chkrs1.Open("SELECT [PARTY].* from [PARTY] where [PARTY].P_CODE='" & pcdd & "' order by [PARTY].P_CODE", xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        Else
            chkrs1.Close()

            MsgBox("Data not exist for this Tenant code .......")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsFileOpen(New FileInfo(Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv")) = True Then
            FileOpen(fnum, Application.StartupPath & "\Reports\Godowndetail.dat", OpenMode.Output)
            FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
            Dim title As String = "Godown Detail"
            globalHeader(title, fnum, fnumm)

            Dim advreceived As Double
            Dim netInvoiceAmt As Double
            Dim recntAMT As Double
            Dim SURVEY, CENSUS, ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GSTt, REMARK As String

            ''''''''''''''''''tenant detail start
            If IsDBNull(chkrs1.Fields(13).Value) Then   'Contact Person
                CPERSON = ""
            Else
                CPERSON = chkrs1.Fields(13).Value
            End If

            If IsDBNull(chkrs1.Fields(18).Value) Then   'Email Address
                EMAIL = ""
            Else
                EMAIL = chkrs1.Fields(18).Value
            End If

            If IsDBNull(chkrs1.Fields(19).Value) Then   'GST
                GSTt = ""
            Else
                GSTt = chkrs1.Fields(19).Value
            End If

            If IsDBNull(chkrs1.Fields(12).Value) Then   'Remark
                REMARK = ""
            Else
                REMARK = chkrs1.Fields(12).Value
            End If

            If IsDBNull(chkrs1.Fields(10).Value) Then    'Address & Phone No
                ADDPHONE = ""
            Else
                ADDPHONE = chkrs1.Fields(10).Value
            End If

            If IsDBNull(chkrs1.Fields(2).Value) Then
                ADD1 = ""
            Else
                ADD1 = chkrs1.Fields(2).Value
            End If

            If IsDBNull(chkrs1.Fields(3).Value) Then
                ADD2 = ""
            Else
                ADD2 = chkrs1.Fields(3).Value
            End If
            If IsDBNull(chkrs1.Fields(4).Value) Then
                ADD3 = ""
            Else
                ADD3 = chkrs1.Fields(4).Value
            End If
            If IsDBNull(chkrs1.Fields(5).Value) Then
                ACT = ""
            Else
                ACT = chkrs1.Fields(5).Value
            End If
            Dim addArr() As String
            If (ADD1.IndexOf(vbLf) >= 0) Then
                addArr = ADD1.Split(vbLf)
                ADD1 = addArr(0)
                If addArr.Length > 1 Then
                    ADD2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                    ADD3 = addArr(2)
                End If
            End If

            If IsDBNull(chkrs1.Fields(11).Value) Then    'House Address & Phone No
                HPHONE = ""
            Else
                HPHONE = chkrs1.Fields(11).Value
            End If
            If IsDBNull(chkrs1.Fields(6).Value) Then
                HADD1 = ""
            Else
                HADD1 = chkrs1.Fields(6).Value

            End If
            If IsDBNull(chkrs1.Fields(7).Value) Then
                HADD2 = ""
            Else
                HADD2 = chkrs1.Fields(7).Value
            End If
            If IsDBNull(chkrs1.Fields(8).Value) Then
                HADD3 = ""
            Else
                HADD3 = chkrs1.Fields(8).Value
            End If
            If IsDBNull(chkrs1.Fields(9).Value) Then
                HCT = ""
            Else
                HCT = chkrs1.Fields(9).Value
            End If
            Dim addHrr() As String
            If (HADD1.IndexOf(vbLf) >= 0) Then
                addHrr = HADD1.Split(vbLf)
                HADD1 = addHrr(0)
                If addHrr.Length > 1 Then
                    HADD2 = addHrr(1)
                End If
                If addHrr.Length > 2 Then
                    HADD3 = addHrr(2)
                End If
            End If
            Print(fnum, GetStringToPrint(20, "Tenant Code       : ", "S") & GetStringToPrint(20, chkrs1.Fields(0).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Tenant Code       :", "S") & "," & GetStringToPrint(20, chkrs1.Fields(0).Value, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Tenant Name       : ", "S") & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Tenant Name       : ", "S") & "," & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Office Address    : ", "S") & GetStringToPrint(55, ADD1, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Office Address    : ", "S") & "," & GetStringToPrint(55, ADD1.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ADD2, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ADD2.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ADD3, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ADD3.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ACT, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ACT, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Office Phone      : ", "S") & GetStringToPrint(55, ADDPHONE, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Office Phone      :", "S") & "," & GetStringToPrint(55, ADDPHONE.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Residence Address : ", "S") & GetStringToPrint(55, HADD1, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Residence Address :", "S") & "," & GetStringToPrint(55, HADD1.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HADD2, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HADD2.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HADD3, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HADD3.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HCT, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HCT.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Residence Phone   : ", "S") & GetStringToPrint(55, HPHONE, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Residence Phone   : ", "S") & "," & GetStringToPrint(55, HPHONE.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Contact Person    : ", "S") & GetStringToPrint(55, CPERSON, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Contact Person    : ", "S") & "," & GetStringToPrint(55, CPERSON.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Email Id          : ", "S") & GetStringToPrint(55, EMAIL, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Email Id          :", "S") & "," & GetStringToPrint(55, EMAIL, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "GSTIN             : ", "S") & GetStringToPrint(55, GSTt, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "GSTIN             : ", "S") & "," & GetStringToPrint(55, GSTt, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "  ", "S") & "," & vbNewLine)
            ''''''''''''''''''tenant detail end

            ''''''''''''''godown detail start
            Dim srs As String = "SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' and [GODOWN].P_CODE='" & pcdd & "' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D"

            chkrs2.Open("SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' and [GODOWN].P_CODE='" & pcdd & "' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D", xcon)

            If chkrs2.BOF = False Then
                chkrs2.MoveFirst()
            End If
            Do While chkrs2.EOF = False
                If chkrs2.Fields(1).Value = pcdd Then

                    Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(20, "Details for Godown :", "S") & GetStringToPrint(55, chkrs2.Fields(0).Value & chkrs2.Fields(3).Value, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "Details for Godown :", "S") & "," & GetStringToPrint(55, chkrs2.Fields(0).Value & chkrs2.Fields(3).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(20, "================== ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "================== ", "S") & vbNewLine)
                    If IsDBNull(chkrs2.Fields(4).Value) Then
                        SURVEY = ""
                    Else
                        SURVEY = chkrs2.Fields(4).Value
                    End If
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                        CENSUS = ""
                    Else
                        CENSUS = chkrs2.Fields(5).Value
                    End If
                    Print(fnum, GetStringToPrint(18, "Survey No.      : ", "S") & GetStringToPrint(20, SURVEY, "S") & GetStringToPrint(12, "Census No.: ", "S") & GetStringToPrint(20, CENSUS, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Survey No.      : ", "S") & "," & GetStringToPrint(20, SURVEY.Replace(",", " "), "S") & "," & GetStringToPrint(12, "Census No.: ", "S") & "," & GetStringToPrint(20, CENSUS.Replace(",", " "), "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(18, "Size            : ", "S") & GetStringToPrint(19, "Length : " & Trim(Format(chkrs2.Fields(19).Value, "##0.00")), "S") & GetStringToPrint(5, " X   ", "S") & GetStringToPrint(15, "Width : " & Trim(Format(chkrs2.Fields(18).Value, "##0.00")), "S") & GetStringToPrint(3, " = ", "S") & GetStringToPrint(25, Trim(Format(chkrs2.Fields(20).Value, "#####0.00")), "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Size            : ", "S") & "," & GetStringToPrint(19, "Length : " & Trim(Format(chkrs2.Fields(19).Value, "##0.00")), "S") & GetStringToPrint(5, " X   ", "S") & GetStringToPrint(15, "Width : " & Trim(Format(chkrs2.Fields(18).Value, "##0.00")), "S") & GetStringToPrint(3, " = ", "S") & GetStringToPrint(25, Trim(Format(chkrs2.Fields(20).Value, "#####0.00")), "S") & vbNewLine)
                    chkrs3.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and P_CODE ='" & pcdd & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    Dim amt, pamt As Double
                    amt = 0
                    pamt = 0
                    If chkrs3.EOF = False Then
                        chkrs3.MoveFirst()

                        amt = chkrs3.Fields(4).Value
                        If IsDBNull(chkrs3.Fields(5).Value) Then
                        Else
                            pamt = chkrs3.Fields(5).Value
                        End If
                    End If
                    chkrs3.Close()
                    Print(fnum, GetStringToPrint(18, "Rent            : ", "S") & GetStringToPrint(12, Format(amt, "#####0.00"), "S") & GetStringToPrint(8, "Prent : ", "S") & GetStringToPrint(12, Format(pamt, "#####0.00"), "S") & GetStringToPrint(8, "Total : ", "S") & GetStringToPrint(15, Format(amt + pamt, "#####0.00"), "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Rent            : ", "S") & "," & GetStringToPrint(12, Format(amt, "#####0.00"), "S") & "," & GetStringToPrint(8, "Prent : ", "S") & "," & GetStringToPrint(12, Format(pamt, "#####0.00"), "S") & "," & GetStringToPrint(8, "Total : ", "S") & "," & GetStringToPrint(15, Format(amt + pamt, "#####0.00"), "S") & vbNewLine)


                    If (chkrs2.Fields(10).Value.Equals("C")) Then
                        Print(fnum, GetStringToPrint(18, "Status          : ", "S") & GetStringToPrint(55, "In Use", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(18, "Status          : ", "S") & "," & GetStringToPrint(55, "In Use", "S") & vbNewLine)
                        '''''''''''''''''''''''godown transfer detail if any data exist in gdtrans table

                        chkrs4.Open("SELECT [GDTRANS].*,[PARTY].P_NAME FROM GDTRANS INNER JOIN PARTY ON [PARTY].P_CODE=[GDTRANS].NP_CODE WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and NP_CODE ='" & pcdd & "' order by  [GROUP],[GODWN_NO],OP_CODE DESC", xcon)
                        If chkrs4.BOF = False Then
                            chkrs4.MoveFirst()
                        End If
                        If chkrs4.EOF = False Then
                            Print(fnum, GetStringToPrint(18, "Previous Tenant : ", "S") & GetStringToPrint(55, chkrs4.Fields(3).Value, "S") & GetStringToPrint(25, "Tenant Code : " & chkrs4.Fields(2).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Previous Tenant : ", "S") & "," & GetStringToPrint(55, chkrs4.Fields(3).Value, "S") & "," & GetStringToPrint(25, "Tenant Code : " & chkrs4.Fields(2).Value, "S") & vbNewLine)
                            'Print(fnum, GetStringToPrint(13, "Transfer Date : ", "S") & GetStringToPrint(55, chkrs4.Fields(6).Value, "S") & vbNewLine)
                            'Print(fnumm, GetStringToPrint(13, "Transfer Date : ", "S") & "," & GetStringToPrint(55, chkrs4.Fields(6).Value, "S") & vbNewLine)
                        End If

                        chkrs4.Close()

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Print(fnum, GetStringToPrint(18, "Using From      : ", "S") & GetStringToPrint(55, chkrs2.Fields(11).Value, "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(18, "Using From      : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(11).Value, "S") & vbNewLine)
                        If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                            Print(fnum, GetStringToPrint(18, "Remarks         : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Remarks         : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                        Else
                            Print(fnum, GetStringToPrint(18, "Remarks         : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Remarks         : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        ''''''''''''''godown detail end

                        ''''''''''''''advance detail start
                        chkrs77.Open("SELECT * from Advances where P_CODE ='" & pcdd & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "' AND [Advances].[GROUP]='" & chkrs2.Fields(0).Value & "'", xcon)
                        If chkrs77.BOF = False Then
                            chkrs77.MoveFirst()
                        End If
                        Dim ED As Date
                        Do While chkrs77.EOF = False

                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Opening Advance Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Opening Advance Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "=======================  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "=======================  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(26, "Advance Received Up to : ", "S") & GetStringToPrint(13, chkrs77.Fields(3).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(26, "Advance Received Up to : ", "S") & "," & GetStringToPrint(13, chkrs77.Fields(3).Value, "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(13, "Receipt No :", "S") & GetStringToPrint(12, chkrs77.Fields(5).Value, "N") & GetStringToPrint(15, "  Receipt Date:", "S") & GetStringToPrint(12, chkrs77.Fields(4).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Receipt No :", "S") & "," & GetStringToPrint(12, chkrs77.Fields(5).Value, "N") & "," & GetStringToPrint(15, "  Receipt Date:", "S") & "," & GetStringToPrint(12, chkrs77.Fields(4).Value, "S") & vbNewLine)
                            ED = chkrs77.Fields(3).Value
                            If chkrs77.EOF = False Then
                                chkrs77.MoveNext()
                            End If
                        Loop
                        chkrs77.Close()
                        ''''''''''''''advance detail end

                        Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)

                        '''''''''''''''''''' invoice details start
                        chkrs66.Open("SELECT [BILL].* from [BILL] where [BILL].BILL_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') and [BILL].BILL_DATE <= FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') and [BILL].P_CODE ='" & pcdd & "' AND [BILL].GODWN_NO='" & chkrs2.Fields(3).Value & "' AND [BILL].[GROUP]='" & chkrs2.Fields(0).Value & "' order by [BILL].BILL_DATE,[BILL].INVOICE_NO", xcon)
                        If chkrs66.BOF = False Then
                            chkrs66.MoveFirst()
                        End If
                        Dim firstInv As Boolean = True
                        Dim totamtt As Double = 0
                        Dim totcgst As Double = 0
                        Dim totsgst As Double = 0
                        Dim totnet As Double = 0
                        Dim advRec As Double = 0


                        Do While chkrs66.EOF = False
                            If firstInv Then

                                Print(fnum, GetStringToPrint(11, "Invoice ", "S") & GetStringToPrint(10, "Invoice", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(7, "  GST", "N") & GetStringToPrint(12, " CGST Amt", "N") & GetStringToPrint(12, " SGST Amt", "N") & GetStringToPrint(12, " Net Amt", "N") & GetStringToPrint(10, "  Against", "S") & GetStringToPrint(10, "  paid/due", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, "Invoice ", "S") & "," & GetStringToPrint(10, "Invoice", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(7, "  GST", "N") & "," & GetStringToPrint(12, " CGST Amt", "N") & "," & GetStringToPrint(12, " SGST Amt", "N") & "," & GetStringToPrint(12, " Net Amt", "N") & "," & GetStringToPrint(10, "  Against", "S") & "," & GetStringToPrint(10, "  paid/due", "S") & vbNewLine)

                                Print(fnum, GetStringToPrint(11, "Date", "S") & GetStringToPrint(10, "No.", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(7, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(10, "  Advance", "S") & GetStringToPrint(10, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, "Date", "S") & "," & GetStringToPrint(10, "No.", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(7, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(10, "  Advance", "S") & "," & GetStringToPrint(10, " ", "S") & vbNewLine)

                                Print(fnum, StrDup(100, "=") & vbNewLine)
                                Print(fnumm, StrDup(100, "=") & vbNewLine)
                                firstInv = False
                                xline = xline + 3
                            End If


                            If (chkrs66.Fields(4).Value >= Format(DateTimePicker2.Value.Date, "dd/MM/yyyy") And chkrs66.Fields(4).Value <= Format(DateTimePicker3.Value.Date, "dd/MM/yyyy")) Then
                                Print(fnum, GetStringToPrint(11, chkrs66.Fields(4).Value, "S") & GetStringToPrint(10, "  " & chkrs66.Fields(0).Value, "S") & GetStringToPrint(13, chkrs66.Fields(5).Value, "N") & GetStringToPrint(7, "   18.0 %", "S") & GetStringToPrint(12, chkrs66.Fields(7).Value, "N") & GetStringToPrint(12, chkrs66.Fields(9).Value, "N") & GetStringToPrint(12, chkrs66.Fields(10).Value, "N") & GetStringToPrint(10, IIf(chkrs66.Fields(15).Value = True, "     Yes", "     No"), "S") & GetStringToPrint(10, IIf(IsDBNull(chkrs66.Fields(13).Value), "     Due", "     Paid"), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, chkrs66.Fields(4).Value, "S") & "," & GetStringToPrint(10, "  " & chkrs66.Fields(0).Value, "S") & "," & GetStringToPrint(13, chkrs66.Fields(5).Value, "N") & "," & GetStringToPrint(7, "   18.0 %", "S") & "," & GetStringToPrint(12, chkrs66.Fields(7).Value, "N") & "," & GetStringToPrint(12, chkrs66.Fields(9).Value, "N") & "," & GetStringToPrint(12, chkrs66.Fields(10).Value, "N") & "," & GetStringToPrint(10, IIf(chkrs66.Fields(15).Value = True, "     Yes", "     No"), "S") & "," & GetStringToPrint(10, IIf(IsDBNull(chkrs66.Fields(13).Value), "     Due", "     Paid"), "S") & vbNewLine)

                                totamtt = totamtt + chkrs66.Fields(5).Value
                                totcgst = totcgst + chkrs66.Fields(7).Value
                                totsgst = totsgst + chkrs66.Fields(9).Value
                                totnet = totnet + chkrs66.Fields(10).Value
                                If chkrs66.Fields(4).Value <= ED Then
                                    advRec = advRec + chkrs66.Fields(10).Value
                                End If
                            End If
                            netInvoiceAmt = netInvoiceAmt + chkrs66.Fields(10).Value
                            If chkrs66.Fields(4).Value <= ED Then
                                advreceived = advreceived + chkrs66.Fields(10).Value
                            End If
                            If chkrs66.EOF = False Then
                                chkrs66.MoveNext()
                            End If
                        Loop

                        chkrs66.Close()
                        Print(fnum, StrDup(100, "=") & vbNewLine)
                        Print(fnumm, StrDup(100, "=") & vbNewLine)
                        Print(fnum, GetStringToPrint(11, "", "S") & GetStringToPrint(10, "", "S") & GetStringToPrint(13, totamtt, "N") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, totcgst, "N") & GetStringToPrint(12, totsgst, "N") & GetStringToPrint(12, totnet, "N") & GetStringToPrint(20, "", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(11, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(13, totamtt, "N") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, totcgst, "N") & "," & GetStringToPrint(12, totsgst, "N") & "," & GetStringToPrint(12, totnet, "N") & "," & GetStringToPrint(20, "", "S") & vbNewLine)
                        '''''''''''''''''''' invoice details end


                        ''''''''''''''''''''payment details start
                        Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"
                        chkrs11.Open("SELECT DISTINCT [receipt].*,[bill].[P_CODE],[BILL].[REC_DATE],[BILL].[REC_NO] from [receipt] INNER JOIN [bill] on [receipt].[rec_no]=int([bill].[rec_no]) and [receipt].[rec_date]=[bill].[rec_date] WHERE [receipt].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [receipt].[GODWN_NO]='" & chkrs2.Fields(3).Value & "' and [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [receipt].REC_DATE <= FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') and [BILL].P_CODE='" & chkrs2.Fields(1).Value & "' order by [receipt].[GROUP],[receipt].GODWN_NO,[receipt].REC_DATE,[receipt].REC_NO", xcon)
                        If chkrs11.BOF = False Then
                            chkrs11.MoveFirst()
                        End If
                        Dim first As Boolean = True
                        Dim totamt As Double = 0
                        Dim totadv As Double = 0
                        Do While chkrs11.EOF = False

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' AND GODWN_NO='" & chkrs11.Fields(2).Value & "' AND [STATUS]='C'", xcon)

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
                            If chkrs44.EOF = False Then
                                If IsDBNull(chkrs44.Fields(5).Value) Then
                                Else
                                    CENSUS = chkrs44.Fields(5).Value
                                End If
                                If IsDBNull(chkrs44.Fields(4).Value) Then
                                Else
                                    SURVEY = chkrs44.Fields(4).Value
                                End If
                                pname = chkrs44.Fields(38).Value
                                pcode1 = chkrs44.Fields(1).Value

                                chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                Dim amtt As Double = 0
                                If chkrs22.EOF = False Then
                                    chkrs22.MoveFirst()
                                    amtt = chkrs22.Fields(4).Value
                                    If IsDBNull(chkrs22.Fields(5).Value) Then
                                    Else
                                        amtt = amtt + chkrs22.Fields(5).Value
                                    End If
                                End If
                                chkrs22.Close()
                                chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                If chkrs33.EOF = False Then
                                    If IsDBNull(chkrs33.Fields(2).Value) Then
                                        CGST_RATE = 0
                                    Else
                                        CGST_RATE = chkrs33.Fields(2).Value
                                    End If
                                    If IsDBNull(chkrs33.Fields(3).Value) Then
                                        SGST_RATE = 0
                                    Else
                                        SGST_RATE = chkrs33.Fields(3).Value
                                    End If
                                End If
                                gst = CGST_RATE + SGST_RATE
                                chkrs33.Close()
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
                            chkrs44.Close()

                            Dim grp As String = chkrs11.Fields(1).Value
                            Dim gdn As String = chkrs11.Fields(2).Value
                            Dim invdt As DateTime = chkrs11.Fields(3).Value
                            Dim inv As Integer = chkrs11.Fields(4).Value
                            Dim FIRSTREC As Boolean = True
                            Dim FROMNO As String = ""
                            Dim TONO As String
                            Dim against As String = ""
                            Dim against1 As String = ""
                            Dim against3 As String = ""
                            Dim against2 As String = ""
                            Dim agcount As Integer = 0
                            Dim adjusted_amt As Double = 0
                            Dim last_bldate As DateTime
                            Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                            chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                            Do While chkrs22.EOF = False
                                If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                    If FIRSTREC Then
                                        chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                        If chkrs6.EOF = False Then
                                            FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                            TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                        Else
                                            FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                            TONO = FROMNO
                                        End If
                                        chkrs6.Close()

                                        FIRSTREC = False
                                    Else
                                        TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                    End If
                                    pname = chkrs22.Fields(16).Value
                                    pcode1 = chkrs22.Fields(3).Value
                                    adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                    last_bldate = chkrs22.Fields(4).Value
                                    If agcount < 7 Then
                                        against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                    Else
                                        If agcount < 14 Then
                                            against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 21 Then
                                                against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 28 Then
                                                    against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                End If
                                            End If
                                        End If
                                    End If

                                    agcount = agcount + 1
                                End If
                                If chkrs22.EOF = False Then
                                    chkrs22.MoveNext()
                                End If
                                If chkrs22.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs22.Close()
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
                            advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                            advanceamtprint = advanceamt
                            If advanceamt > 0 Then
                                Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs55.Open(Rss, xcon)
                                Do While chkrs55.EOF = False
                                    If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                        lastbilladjusted = chkrs55.Fields(0).Value
                                        last_bldate = chkrs55.Fields(4).Value
                                    End If
                                    If chkrs55.EOF = False Then
                                        chkrs55.MoveNext()
                                    End If
                                Loop
                                chkrs55.Close()
                                If lastbilladjusted = 0 Then
                                    Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND P_CODE='" & pcode1 & "' and [STATUS]='C' order by [GROUP]+GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    If chkrs55.EOF = False Then
                                        last_bldate = chkrs55.Fields(0).Value
                                    End If
                                    chkrs55.Close()
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
                            recntAMT = recntAMT + chkrs11.Fields(5).Value
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If (chkrs11.Fields(3).Value >= Format(DateTimePicker5.Value.Date, "dd/MM/yyyy") And chkrs11.Fields(3).Value <= Format(DateTimePicker4.Value.Date, "dd/MM/yyyy")) Then
                                If first Then
                                    Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(50, "Tenant Name", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & vbNewLine)

                                    Print(fnum, GetStringToPrint(17, "Godown Type", "S") & GetStringToPrint(17, "Godown No.", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "  ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(50, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Godown Type", "S") & "," & GetStringToPrint(17, "Godown No.", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(50, "Bank A/C Detail", "S") & vbNewLine)



                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, "Bank Name & Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, "Bank Name & Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(133, "=") & vbNewLine)
                                    Print(fnumm, StrDup(133, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                Else
                                    Print(fnum, StrDup(30, ">") & vbNewLine)
                                End If


                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(50, pname, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(50, pname, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(1).Value, "S") & GetStringToPrint(17, chkrs11.Fields(2).Value.ToString, "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(17, chkrs11.Fields(2).Value.ToString, "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.Fields(7).Value.Equals("C") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(50, chkrs11.Fields(8).Value & " " & chkrs11.Fields(9).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, chkrs11.Fields(8).Value & " " & chkrs11.Fields(9).Value, "S") & vbNewLine)
                                    xline = xline + 1
                                End If

                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                xline = xline + 2

                            End If

                            If chkrs11.EOF = False Then
                                chkrs11.MoveNext()
                            End If
                            If chkrs11.EOF = True Then
                                Exit Do
                            End If

                        Loop
                        If totamt > 0 Then
                            Print(fnum, StrDup(133, "-") & vbNewLine)
                            Print(fnumm, StrDup(133, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(50, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(50, "", "S") & vbNewLine)
                        End If
                        chkrs11.Close()

                        ''''''''''''''''''''payment details end
                    Else
                        ''''''''''''''''checking in godown transfer table
                        chkrs4.Open("SELECT [GDTRANS].*,[PARTY].P_NAME FROM GDTRANS INNER JOIN PARTY ON [PARTY].P_CODE=[GDTRANS].NP_CODE WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and OP_CODE ='" & TextBox1.Text & "' order by  [GROUP],[GODWN_NO],OP_CODE DESC", xcon)
                        If chkrs4.BOF = False Then
                            chkrs4.MoveFirst()
                        End If
                        Do While chkrs4.EOF = False
                            Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Transfered to ", "S") & GetStringToPrint(55, chkrs4.Fields(7).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Transfered to ", "S") & "," & GetStringToPrint(55, chkrs4.Fields(7).Value, "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(13, "Used From  : ", "S") & GetStringToPrint(12, chkrs2.Fields(11).Value, "S") & GetStringToPrint(9, "up to :", "S") & GetStringToPrint(15, chkrs4.Fields(6).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Used From  : ", "S") & "," & GetStringToPrint(12, chkrs2.Fields(11).Value, "S") & "," & GetStringToPrint(9, "up to :", "S") & "," & GetStringToPrint(15, chkrs4.Fields(6).Value, "S") & vbNewLine)
                            If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            End If
                            ''''''''''''''''''''payment details
                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Dim STR As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs4.Fields(6).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"
                            chkrs11.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs4.Fields(6).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
                            If chkrs11.BOF = False Then
                                chkrs11.MoveFirst()
                            End If
                            Dim first As Boolean = True
                            Dim totamt As Double = 0
                            Dim totadv As Double = 0
                            Do While chkrs11.EOF = False

                                If first Then

                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(180, "=") & vbNewLine)
                                    Print(fnumm, StrDup(180, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                End If
                                Dim srn As String = "SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE P_CODE='" & TextBox1.Text & "' AND [GROUP]='" & chkrs2.Fields(0).Value & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "'"
                                chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GODOWN].P_CODE='" & TextBox1.Text & "' AND [GROUP]='" & chkrs2.Fields(0).Value & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "'", xcon)

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
                                If chkrs44.EOF = False Then
                                    If IsDBNull(chkrs44.Fields(5).Value) Then
                                    Else
                                        CENSUS = chkrs44.Fields(5).Value
                                    End If
                                    If IsDBNull(chkrs44.Fields(4).Value) Then
                                    Else
                                        SURVEY = chkrs44.Fields(4).Value
                                    End If
                                    pname = chkrs44.Fields(38).Value
                                    pcode1 = TextBox1.Text

                                    chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                    Dim amtt As Double = 0
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveFirst()
                                        amtt = chkrs22.Fields(4).Value
                                        If IsDBNull(chkrs22.Fields(5).Value) Then
                                        Else
                                            amtt = amtt + chkrs22.Fields(5).Value
                                        End If
                                    End If
                                    chkrs22.Close()
                                    chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                    If chkrs33.EOF = False Then
                                        If IsDBNull(chkrs33.Fields(2).Value) Then
                                            CGST_RATE = 0
                                        Else
                                            CGST_RATE = chkrs33.Fields(2).Value
                                        End If
                                        If IsDBNull(chkrs33.Fields(3).Value) Then
                                            SGST_RATE = 0
                                        Else
                                            SGST_RATE = chkrs33.Fields(3).Value
                                        End If
                                    End If
                                    gst = CGST_RATE + SGST_RATE
                                    chkrs33.Close()
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
                                chkrs44.Close()

                                Dim grp As String = chkrs11.Fields(1).Value
                                Dim gdn As String = chkrs11.Fields(2).Value
                                Dim invdt As DateTime = chkrs11.Fields(3).Value
                                Dim inv As Integer = chkrs11.Fields(4).Value
                                Dim FIRSTREC As Boolean = True
                                Dim FROMNO As String = ""
                                Dim TONO As String
                                Dim against As String = ""
                                Dim against1 As String = ""
                                Dim against3 As String = ""
                                Dim against2 As String = ""
                                Dim agcount As Integer = 0
                                Dim adjusted_amt As Double = 0
                                Dim last_bldate As DateTime
                                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                                Do While chkrs22.EOF = False
                                    If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                        If FIRSTREC Then
                                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                            If chkrs6.EOF = False Then
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                            Else
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                                TONO = FROMNO
                                            End If
                                            chkrs6.Close()

                                            FIRSTREC = False
                                        Else
                                            TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                        End If
                                        pname = chkrs22.Fields(16).Value
                                        pcode1 = chkrs22.Fields(3).Value
                                        adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                        last_bldate = chkrs22.Fields(4).Value
                                        If agcount < 7 Then
                                            against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 14 Then
                                                against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 21 Then
                                                    against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                    If agcount < 28 Then
                                                        against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If

                                        agcount = agcount + 1
                                    End If
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveNext()
                                    End If
                                    If chkrs22.EOF = True Then
                                        Exit Do
                                    End If

                                Loop
                                chkrs22.Close()
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
                                advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                                advanceamtprint = advanceamt
                                If advanceamt > 0 Then
                                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    Do While chkrs55.EOF = False
                                        If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                            lastbilladjusted = chkrs55.Fields(0).Value
                                            last_bldate = chkrs55.Fields(4).Value
                                        End If
                                        If chkrs55.EOF = False Then
                                            chkrs55.MoveNext()
                                        End If
                                    Loop
                                    chkrs55.Close()
                                    If lastbilladjusted = 0 Then
                                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' and  P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                                        chkrs55.Open(Rss, xcon)
                                        If chkrs55.EOF = False Then
                                            last_bldate = chkrs55.Fields(0).Value
                                        End If
                                        chkrs55.Close()
                                    End If
                                    Dim dtcounter As Long = 1
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

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.EOF = False Then
                                    chkrs11.MoveNext()
                                End If
                                If chkrs11.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs11.Close()
                            Print(fnum, StrDup(180, "-") & vbNewLine)
                            Print(fnumm, StrDup(180, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                            ''''''''''''''''''''payment details
                            If chkrs4.BOF = False Then
                                chkrs4.MoveNext()
                            End If
                        Loop
                        chkrs4.Close()
                        ''''''''''''''''''''details of payment for godown


                        '''''''''''''''''''checking in godown close /suspend tabel
                        chkrs5.Open("SELECT * FROM CLGDWN WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and P_CODE ='" & TextBox1.Text & "' order by [GROUP],GODWN_NO,P_CODE DESC", xcon)
                        If chkrs5.BOF = False Then
                            chkrs5.MoveFirst()
                        End If
                        Do While chkrs5.EOF = False
                            If chkrs5.Fields(6).Value.Equals("S") Then
                                Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Suspended", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Suspended", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Closed", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Closed", "S") & vbNewLine)
                            End If

                            Print(fnum, GetStringToPrint(13, "Used From  : ", "S") & GetStringToPrint(15, chkrs2.Fields(11).Value, "S") & GetStringToPrint(13, "up to  : ", "S") & GetStringToPrint(15, chkrs5.Fields(3).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Used From  : ", "S") & "," & GetStringToPrint(15, chkrs2.Fields(11).Value, "S") & "," & GetStringToPrint(13, "up to  : ", "S") & "," & GetStringToPrint(15, chkrs5.Fields(3).Value, "S") & vbNewLine)
                            If IsDBNull(chkrs5.Fields(5).Value) Or chkrs5.Fields(5).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Reason     : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Reason     : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Reason     : ", "S") & GetStringToPrint(55, chkrs5.Fields(5).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Reason     : ", "S") & "," & GetStringToPrint(55, chkrs5.Fields(5).Value, "S") & vbNewLine)
                            End If

                            If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            End If
                            ''''''''''''''''''''payment details
                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)

                            chkrs11.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs5.Fields(3).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
                            If chkrs11.BOF = False Then
                                chkrs11.MoveFirst()
                            End If
                            Dim first As Boolean = True
                            Dim totamt As Double = 0
                            Dim totadv As Double = 0
                            Do While chkrs11.EOF = False

                                If first Then

                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(180, "=") & vbNewLine)
                                    Print(fnumm, StrDup(180, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                End If

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)

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
                                If chkrs44.EOF = False Then
                                    If IsDBNull(chkrs44.Fields(5).Value) Then
                                    Else
                                        CENSUS = chkrs44.Fields(5).Value
                                    End If
                                    If IsDBNull(chkrs44.Fields(4).Value) Then
                                    Else
                                        SURVEY = chkrs44.Fields(4).Value
                                    End If
                                    pname = chkrs44.Fields(38).Value
                                    pcode1 = TextBox1.Text

                                    chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                    Dim amtt As Double = 0
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveFirst()
                                        amtt = chkrs22.Fields(4).Value
                                        If IsDBNull(chkrs22.Fields(5).Value) Then
                                        Else
                                            amtt = amtt + chkrs22.Fields(5).Value
                                        End If
                                    End If
                                    chkrs22.Close()
                                    chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                    If chkrs33.EOF = False Then
                                        If IsDBNull(chkrs33.Fields(2).Value) Then
                                            CGST_RATE = 0
                                        Else
                                            CGST_RATE = chkrs33.Fields(2).Value
                                        End If
                                        If IsDBNull(chkrs33.Fields(3).Value) Then
                                            SGST_RATE = 0
                                        Else
                                            SGST_RATE = chkrs33.Fields(3).Value
                                        End If
                                    End If
                                    gst = CGST_RATE + SGST_RATE
                                    chkrs33.Close()
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
                                chkrs44.Close()

                                Dim grp As String = chkrs11.Fields(1).Value
                                Dim gdn As String = chkrs11.Fields(2).Value
                                Dim invdt As DateTime = chkrs11.Fields(3).Value
                                Dim inv As Integer = chkrs11.Fields(4).Value
                                Dim FIRSTREC As Boolean = True
                                Dim FROMNO As String = ""
                                Dim TONO As String
                                Dim against As String = ""
                                Dim against1 As String = ""
                                Dim against3 As String = ""
                                Dim against2 As String = ""
                                Dim agcount As Integer = 0
                                Dim adjusted_amt As Double = 0
                                Dim last_bldate As DateTime
                                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                                Do While chkrs22.EOF = False
                                    If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                        If FIRSTREC Then
                                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                            If chkrs6.EOF = False Then
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                            Else
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                                TONO = FROMNO
                                            End If
                                            chkrs6.Close()

                                            FIRSTREC = False
                                        Else
                                            TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                        End If
                                        pname = chkrs22.Fields(16).Value
                                        pcode1 = chkrs22.Fields(3).Value
                                        adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                        last_bldate = chkrs22.Fields(4).Value
                                        If agcount < 7 Then
                                            against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 14 Then
                                                against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 21 Then
                                                    against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                    If agcount < 28 Then
                                                        against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If

                                        agcount = agcount + 1
                                    End If
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveNext()
                                    End If
                                    If chkrs22.EOF = True Then
                                        Exit Do
                                    End If

                                Loop
                                chkrs22.Close()
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

                                advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                                advanceamtprint = advanceamt
                                If advanceamt > 0 Then
                                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    Do While chkrs55.EOF = False
                                        If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                            lastbilladjusted = chkrs55.Fields(0).Value
                                            last_bldate = chkrs55.Fields(4).Value
                                        End If
                                        If chkrs55.EOF = False Then
                                            chkrs55.MoveNext()
                                        End If
                                    Loop
                                    chkrs55.Close()
                                    If lastbilladjusted = 0 Then
                                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND [STATUS]='C' AND P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                                        chkrs55.Open(Rss, xcon)
                                        If chkrs55.EOF = False Then
                                            last_bldate = chkrs55.Fields(0).Value
                                        End If
                                        chkrs55.Close()
                                    End If
                                    Dim dtcounter As Long = 1
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

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.EOF = False Then
                                    chkrs11.MoveNext()
                                End If
                                If chkrs11.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs11.Close()
                            Print(fnum, StrDup(180, "-") & vbNewLine)
                            Print(fnumm, StrDup(180, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                            ''''''''''''''''''''payment details
                            If chkrs5.BOF = False Then
                                chkrs5.MoveNext()
                            End If
                        Loop
                        chkrs5.Close()

                    End If


                End If
                Dim otst As Double = 0
                Dim prvdt As String = ""
                If IsDBNull(chkrs2.Fields(21).Value) Then   '''''outst from godown table  - due amount
                Else
                    otst = chkrs2.Fields(21).Value
                End If

                prvdt = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToInt16(chkrs2.Fields(14).Value)).ToString() + " - " + chkrs2.Fields(15).Value.ToString()
                If chkrs2.BOF = False Then
                    chkrs2.MoveNext()
                End If
                Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                If (otst > 0) Then
                    Print(fnum, GetStringToPrint(42, "Previous Outstanding Detail as on Date : ", "S") & GetStringToPrint(15, Format(otst, "########0.00"), "N") & GetStringToPrint(30, " from " & prvdt, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(42, "Previous Outstanding Detail as on Date : ", "S") & GetStringToPrint(15, Format(otst, "########0.00"), "N") & GetStringToPrint(30, " from " & prvdt, "S") & vbNewLine)
                End If
                Print(fnum, GetStringToPrint(42, "Outstanding Detail as on Date : " + DateTimePicker1.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(42, "Outstanding Detail as on Date : " + DateTimePicker1.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(42, "===============================================  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(42, "===============================================  ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Advance Opening -->  " & advreceived, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Advance Opening -->  " & advreceived, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Receipt tot     -->  " & recntAMT, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Receipt tot     -->  " & recntAMT, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Invoice Amt     -->  " & netInvoiceAmt, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Invoice Amt     -->  " & netInvoiceAmt, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "===================  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "===================  ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(21, "Outstanding Amt -->  ", "S"))
                Print(fnumm, GetStringToPrint(21, "Outstanding Amt -->  ", "S"))
                Print(fnum, GetStringToPrint(30, IIf((netInvoiceAmt - recntAMT - advreceived) > 0, (netInvoiceAmt - recntAMT - advreceived), ((netInvoiceAmt - recntAMT - advreceived) * -1) & " Advance"), "S"))
                Print(fnumm, GetStringToPrint(30, IIf((netInvoiceAmt - recntAMT - advreceived) > 0, (netInvoiceAmt - recntAMT - advreceived), ((netInvoiceAmt - recntAMT - advreceived) * -1) & " Advance"), "S"))
            Loop
            chkrs2.Close()
            chkrs1.Close()

            FileClose(fnum)
            FileClose(fnumm)
            MyConn.Close()
            FrmGdnDtlView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Godowndetail.dat", RichTextBoxStreamType.PlainText)

            FrmGdnDtlView.Show()
            CreatePDF(Application.StartupPath & "\Reports\Godowndetail.dat", Application.StartupPath & "\Reports\Godowndetail")
            MsgBox(Application.StartupPath + " \Reports\" & TextBox5.Text & ".csv file is generated")
        End If
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        ''''''''create pdf from .bat file
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 10

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"

            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 15

            pdfPage.Width = 842
            pdfPage.Height = 595
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER New", 9, XFontStyle.Regular)

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
                        font = New XFont("COURIER New", 9, XFontStyle.Regular)

                        pdfPage.TrimMargins.Left = 15

                        pdfPage.Width = 842
                        pdfPage.Height = 595
                        yPoint = 10
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
    Public Function IsFileOpen(ByVal file As FileInfo)
        'Dim stream As FileStream = Nothing
        'Try
        '    stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
        '    stream.Close()
        Return True
        'Catch ex As Exception

        '    If TypeOf ex Is IOException Then
        '        MsgBox("Please close file " + file.FullName)
        '        Return False

        '        ' do something here, either close the file if you have a handle, show a msgbox, retry  or as a last resort terminate the process - which could cause corruption and lose data
        '    End If
        'End Try

    End Function

    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        ''''''assign godwn_no value from current datagrid row to godown number text box and group to godown group combobox when double click the datagrid
        TextBox1.Text = GetValue(DataGridView2.Item(3, DataGridView2.CurrentRow.Index).Value)
        ComboBox1.Text = GetValue(DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value)
        ComboBox1.SelectedIndex = ComboBox1.FindStringExact(ComboBox1.Text)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()     ''''''close form
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '''''''''print report , same logic as button1_click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter Tenant code .......")
            TextBox1.Focus()
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        chkrs222.Open("SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' AND [GODOWN].[STATUS]='C' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D", xcon)

        If chkrs222.BOF = False Then
            chkrs222.MoveFirst()
        End If
        Dim pcdd As String = ""
        Do While chkrs222.EOF = False
            pcdd = chkrs222.Fields(1).Value
            Exit Do
        Loop
        chkrs222.Close()
        chkrs1.Open("SELECT [PARTY].* from [PARTY] where [PARTY].P_CODE='" & pcdd & "' order by [PARTY].P_CODE", xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        Else
            chkrs1.Close()

            MsgBox("Data not exist for this Tenant code .......")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsFileOpen(New FileInfo(Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv")) = True Then
            FileOpen(fnum, Application.StartupPath & "\Reports\Godowndetail.dat", OpenMode.Output)
            FileOpen(fnumm, Application.StartupPath & "\Reports\" & TextBox5.Text & ".csv", OpenMode.Output)
            Dim title As String = "Godown Detail"
            globalHeader(title, fnum, fnumm)

            Dim advreceived As Double
            Dim netInvoiceAmt As Double
            Dim recntAMT As Double
            Dim SURVEY, CENSUS, ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GSTt, REMARK As String
            If IsDBNull(chkrs1.Fields(13).Value) Then   'Contact Person
                CPERSON = ""
            Else
                CPERSON = chkrs1.Fields(13).Value
            End If

            If IsDBNull(chkrs1.Fields(18).Value) Then   'Email Address
                EMAIL = ""
            Else
                EMAIL = chkrs1.Fields(18).Value
            End If

            If IsDBNull(chkrs1.Fields(19).Value) Then   'GST
                GSTt = ""
            Else
                GSTt = chkrs1.Fields(19).Value
            End If

            If IsDBNull(chkrs1.Fields(12).Value) Then   'Remark
                REMARK = ""
            Else
                REMARK = chkrs1.Fields(12).Value
            End If

            If IsDBNull(chkrs1.Fields(10).Value) Then    'Address & Phone No
                ADDPHONE = ""
            Else
                ADDPHONE = chkrs1.Fields(10).Value
            End If

            If IsDBNull(chkrs1.Fields(2).Value) Then
                ADD1 = ""
            Else
                ADD1 = chkrs1.Fields(2).Value
            End If

            If IsDBNull(chkrs1.Fields(3).Value) Then
                ADD2 = ""
            Else
                ADD2 = chkrs1.Fields(3).Value
            End If
            If IsDBNull(chkrs1.Fields(4).Value) Then
                ADD3 = ""
            Else
                ADD3 = chkrs1.Fields(4).Value
            End If
            If IsDBNull(chkrs1.Fields(5).Value) Then
                ACT = ""
            Else
                ACT = chkrs1.Fields(5).Value
            End If
            Dim addArr() As String
            If (ADD1.IndexOf(vbLf) >= 0) Then
                addArr = ADD1.Split(vbLf)
                ADD1 = addArr(0)
                If addArr.Length > 1 Then
                    ADD2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                    ADD3 = addArr(2)
                End If
            End If

            If IsDBNull(chkrs1.Fields(11).Value) Then    'House Address & Phone No
                HPHONE = ""
            Else
                HPHONE = chkrs1.Fields(11).Value
            End If
            If IsDBNull(chkrs1.Fields(6).Value) Then
                HADD1 = ""
            Else
                HADD1 = chkrs1.Fields(6).Value

            End If
            If IsDBNull(chkrs1.Fields(7).Value) Then
                HADD2 = ""
            Else
                HADD2 = chkrs1.Fields(7).Value
            End If
            If IsDBNull(chkrs1.Fields(8).Value) Then
                HADD3 = ""
            Else
                HADD3 = chkrs1.Fields(8).Value
            End If
            If IsDBNull(chkrs1.Fields(9).Value) Then
                HCT = ""
            Else
                HCT = chkrs1.Fields(9).Value
            End If
            Dim addHrr() As String
            If (HADD1.IndexOf(vbLf) >= 0) Then
                addHrr = HADD1.Split(vbLf)
                HADD1 = addHrr(0)
                If addHrr.Length > 1 Then
                    HADD2 = addHrr(1)
                End If
                If addHrr.Length > 2 Then
                    HADD3 = addHrr(2)
                End If
            End If
            Print(fnum, GetStringToPrint(20, "Tenant Code       : ", "S") & GetStringToPrint(20, chkrs1.Fields(0).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Tenant Code       :", "S") & "," & GetStringToPrint(20, chkrs1.Fields(0).Value, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Tenant Name       : ", "S") & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Tenant Name       : ", "S") & "," & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Office Address    : ", "S") & GetStringToPrint(55, ADD1, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Office Address    : ", "S") & "," & GetStringToPrint(55, ADD1.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ADD2, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ADD2.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ADD3, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ADD3.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, ACT, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, ACT, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Office Phone      : ", "S") & GetStringToPrint(55, ADDPHONE, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Office Phone      :", "S") & "," & GetStringToPrint(55, ADDPHONE.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Residence Address : ", "S") & GetStringToPrint(55, HADD1, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Residence Address :", "S") & "," & GetStringToPrint(55, HADD1.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HADD2, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HADD2.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HADD3, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HADD3.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "", "S") & GetStringToPrint(55, HCT, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "", "S") & "," & GetStringToPrint(55, HCT.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Residence Phone   : ", "S") & GetStringToPrint(55, HPHONE, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Residence Phone   : ", "S") & "," & GetStringToPrint(55, HPHONE.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Contact Person    : ", "S") & GetStringToPrint(55, CPERSON, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Contact Person    : ", "S") & "," & GetStringToPrint(55, CPERSON.Replace(",", " "), "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "Email Id          : ", "S") & GetStringToPrint(55, EMAIL, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "Email Id          :", "S") & "," & GetStringToPrint(55, EMAIL, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "GSTIN             : ", "S") & GetStringToPrint(55, GSTt, "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "GSTIN             : ", "S") & "," & GetStringToPrint(55, GSTt, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
            Print(fnumm, GetStringToPrint(20, "  ", "S") & "," & vbNewLine)
            Dim srs As String = "SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' and [GODOWN].P_CODE='" & pcdd & "' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D"

            chkrs2.Open("SELECT [GODOWN].* from [GODOWN] where [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND [GODOWN].GODWN_NO='" & TextBox1.Text & "' and [GODOWN].P_CODE='" & pcdd & "' AND [GODOWN].[FROM_D]<=FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') order by [GODOWN].[GROUP],[GODOWN].GODWN_NO,[GODOWN].FROM_D", xcon)

            If chkrs2.BOF = False Then
                chkrs2.MoveFirst()
            End If
            Do While chkrs2.EOF = False
                If chkrs2.Fields(1).Value = pcdd Then

                    Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(20, "Details for Godown :", "S") & GetStringToPrint(55, chkrs2.Fields(0).Value & chkrs2.Fields(3).Value, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "Details for Godown :", "S") & "," & GetStringToPrint(55, chkrs2.Fields(0).Value & chkrs2.Fields(3).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(20, "================== ", "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(20, "================== ", "S") & vbNewLine)
                    If IsDBNull(chkrs2.Fields(4).Value) Then
                        SURVEY = ""
                    Else
                        SURVEY = chkrs2.Fields(4).Value
                    End If
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                        CENSUS = ""
                    Else
                        CENSUS = chkrs2.Fields(5).Value
                    End If
                    Print(fnum, GetStringToPrint(18, "Survey No.      : ", "S") & GetStringToPrint(20, SURVEY, "S") & GetStringToPrint(12, "Census No.: ", "S") & GetStringToPrint(20, CENSUS, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Survey No.      : ", "S") & "," & GetStringToPrint(20, SURVEY.Replace(",", " "), "S") & "," & GetStringToPrint(12, "Census No.: ", "S") & "," & GetStringToPrint(20, CENSUS.Replace(",", " "), "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(18, "Size            : ", "S") & GetStringToPrint(19, "Length : " & Trim(Format(chkrs2.Fields(19).Value, "##0.00")), "S") & GetStringToPrint(5, " X   ", "S") & GetStringToPrint(15, "Width : " & Trim(Format(chkrs2.Fields(18).Value, "##0.00")), "S") & GetStringToPrint(3, " = ", "S") & GetStringToPrint(25, Trim(Format(chkrs2.Fields(20).Value, "#####0.00")), "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Size            : ", "S") & "," & GetStringToPrint(19, "Length : " & Trim(Format(chkrs2.Fields(19).Value, "##0.00")), "S") & GetStringToPrint(5, " X   ", "S") & GetStringToPrint(15, "Width : " & Trim(Format(chkrs2.Fields(18).Value, "##0.00")), "S") & GetStringToPrint(3, " = ", "S") & GetStringToPrint(25, Trim(Format(chkrs2.Fields(20).Value, "#####0.00")), "S") & vbNewLine)
                    chkrs3.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and P_CODE ='" & pcdd & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    Dim amt, pamt As Double
                    amt = 0
                    pamt = 0
                    If chkrs3.EOF = False Then
                        chkrs3.MoveFirst()

                        amt = chkrs3.Fields(4).Value
                        If IsDBNull(chkrs3.Fields(5).Value) Then
                        Else
                            pamt = chkrs3.Fields(5).Value
                        End If
                    End If
                    chkrs3.Close()
                    Print(fnum, GetStringToPrint(18, "Rent            : ", "S") & GetStringToPrint(12, Format(amt, "#####0.00"), "S") & GetStringToPrint(8, "Prent : ", "S") & GetStringToPrint(12, Format(pamt, "#####0.00"), "S") & GetStringToPrint(8, "Total : ", "S") & GetStringToPrint(15, Format(amt + pamt, "#####0.00"), "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(18, "Rent            : ", "S") & "," & GetStringToPrint(12, Format(amt, "#####0.00"), "S") & "," & GetStringToPrint(8, "Prent : ", "S") & "," & GetStringToPrint(12, Format(pamt, "#####0.00"), "S") & "," & GetStringToPrint(8, "Total : ", "S") & "," & GetStringToPrint(15, Format(amt + pamt, "#####0.00"), "S") & vbNewLine)

                    If (chkrs2.Fields(10).Value.Equals("C")) Then
                        Print(fnum, GetStringToPrint(18, "Status          : ", "S") & GetStringToPrint(55, "In Use", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(18, "Status          : ", "S") & "," & GetStringToPrint(55, "In Use", "S") & vbNewLine)
                        '''''''''''''''''''''''godown transfer detail if any data exist in gdtrans table

                        chkrs4.Open("SELECT [GDTRANS].*,[PARTY].P_NAME FROM GDTRANS INNER JOIN PARTY ON [PARTY].P_CODE=[GDTRANS].NP_CODE WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and NP_CODE ='" & pcdd & "' order by  [GROUP],[GODWN_NO],OP_CODE DESC", xcon)
                        If chkrs4.BOF = False Then
                            chkrs4.MoveFirst()
                        End If
                        If chkrs4.EOF = False Then
                            Print(fnum, GetStringToPrint(18, "Previous Tenant : ", "S") & GetStringToPrint(55, chkrs4.Fields(3).Value, "S") & GetStringToPrint(25, "Tenant Code : " & chkrs4.Fields(2).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Previous Tenant : ", "S") & "," & GetStringToPrint(55, chkrs4.Fields(3).Value, "S") & "," & GetStringToPrint(25, "Tenant Code : " & chkrs4.Fields(2).Value, "S") & vbNewLine)
                        End If

                        chkrs4.Close()

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Print(fnum, GetStringToPrint(18, "Using From      : ", "S") & GetStringToPrint(55, chkrs2.Fields(11).Value, "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(18, "Using From      : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(11).Value, "S") & vbNewLine)
                        If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                            Print(fnum, GetStringToPrint(18, "Remarks         : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Remarks         : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                        Else
                            Print(fnum, GetStringToPrint(18, "Remarks         : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, "Remarks         : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value.Replace(",", " "), "S") & vbNewLine)
                        End If
                        If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                        Else
                            Print(fnum, GetStringToPrint(18, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(18, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value.Replace(",", " "), "S") & vbNewLine)
                        End If


                        ''''''''''''''advance detail
                        '''''''''''''''''''' invoice details
                        chkrs77.Open("SELECT * from Advances where P_CODE ='" & pcdd & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "' AND [Advances].[GROUP]='" & chkrs2.Fields(0).Value & "'", xcon)
                        If chkrs77.BOF = False Then
                            chkrs77.MoveFirst()
                        End If
                        Dim ED As Date
                        Do While chkrs77.EOF = False

                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Opening Advance Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Opening Advance Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "=======================  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "=======================  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(26, "Advance Received Up to : ", "S") & GetStringToPrint(13, chkrs77.Fields(3).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(26, "Advance Received Up to : ", "S") & "," & GetStringToPrint(13, chkrs77.Fields(3).Value, "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(13, "Receipt No :", "S") & GetStringToPrint(12, chkrs77.Fields(5).Value, "N") & GetStringToPrint(15, "  Receipt Date:", "S") & GetStringToPrint(12, chkrs77.Fields(4).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Receipt No :", "S") & "," & GetStringToPrint(12, chkrs77.Fields(5).Value, "N") & "," & GetStringToPrint(15, "  Receipt Date:", "S") & "," & GetStringToPrint(12, chkrs77.Fields(4).Value, "S") & vbNewLine)
                            ED = chkrs77.Fields(3).Value
                            If chkrs77.EOF = False Then
                                chkrs77.MoveNext()
                            End If
                        Loop
                        chkrs77.Close()
                        Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                        '''''''''''''''''''' invoice details
                        chkrs66.Open("SELECT [BILL].* from [BILL] where [BILL].BILL_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') and [BILL].BILL_DATE <= FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') and [BILL].P_CODE ='" & pcdd & "' AND [BILL].GODWN_NO='" & chkrs2.Fields(3).Value & "' AND [BILL].[GROUP]='" & chkrs2.Fields(0).Value & "' order by [BILL].BILL_DATE,[BILL].INVOICE_NO", xcon)
                        If chkrs66.BOF = False Then
                            chkrs66.MoveFirst()
                        End If
                        Dim firstInv As Boolean = True
                        Dim totamtt As Double = 0
                        Dim totcgst As Double = 0
                        Dim totsgst As Double = 0
                        Dim totnet As Double = 0
                        Dim advRec As Double = 0


                        Do While chkrs66.EOF = False
                            If firstInv Then

                                Print(fnum, GetStringToPrint(11, "Invoice ", "S") & GetStringToPrint(10, "Invoice", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(7, "  GST", "N") & GetStringToPrint(12, " CGST Amt", "N") & GetStringToPrint(12, " SGST Amt", "N") & GetStringToPrint(12, " Net Amt", "N") & GetStringToPrint(10, "  Against", "S") & GetStringToPrint(10, "  paid/due", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, "Invoice ", "S") & "," & GetStringToPrint(10, "Invoice", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(7, "  GST", "N") & "," & GetStringToPrint(12, " CGST Amt", "N") & "," & GetStringToPrint(12, " SGST Amt", "N") & "," & GetStringToPrint(12, " Net Amt", "N") & "," & GetStringToPrint(10, "  Against", "S") & "," & GetStringToPrint(10, "  paid/due", "S") & vbNewLine)

                                Print(fnum, GetStringToPrint(11, "Date", "S") & GetStringToPrint(10, "No.", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(7, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(12, " ", "N") & GetStringToPrint(10, "  Advance", "S") & GetStringToPrint(10, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, "Date", "S") & "," & GetStringToPrint(10, "No.", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(7, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(12, " ", "N") & "," & GetStringToPrint(10, "  Advance", "S") & "," & GetStringToPrint(10, " ", "S") & vbNewLine)

                                Print(fnum, StrDup(100, "=") & vbNewLine)
                                Print(fnumm, StrDup(100, "=") & vbNewLine)
                                firstInv = False
                                xline = xline + 3
                            End If


                            If (chkrs66.Fields(4).Value >= Format(DateTimePicker2.Value.Date, "dd/MM/yyyy") And chkrs66.Fields(4).Value <= Format(DateTimePicker3.Value.Date, "dd/MM/yyyy")) Then
                                Print(fnum, GetStringToPrint(11, chkrs66.Fields(4).Value, "S") & GetStringToPrint(10, "  " & chkrs66.Fields(0).Value, "S") & GetStringToPrint(13, chkrs66.Fields(5).Value, "N") & GetStringToPrint(7, "   18.0 %", "S") & GetStringToPrint(12, chkrs66.Fields(7).Value, "N") & GetStringToPrint(12, chkrs66.Fields(9).Value, "N") & GetStringToPrint(12, chkrs66.Fields(10).Value, "N") & GetStringToPrint(10, IIf(chkrs66.Fields(15).Value = True, "     Yes", "     No"), "S") & GetStringToPrint(10, IIf(IsDBNull(chkrs66.Fields(13).Value), "     Due", "     Paid"), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(11, chkrs66.Fields(4).Value, "S") & "," & GetStringToPrint(10, "  " & chkrs66.Fields(0).Value, "S") & "," & GetStringToPrint(13, chkrs66.Fields(5).Value, "N") & "," & GetStringToPrint(7, "   18.0 %", "S") & "," & GetStringToPrint(12, chkrs66.Fields(7).Value, "N") & "," & GetStringToPrint(12, chkrs66.Fields(9).Value, "N") & "," & GetStringToPrint(12, chkrs66.Fields(10).Value, "N") & "," & GetStringToPrint(10, IIf(chkrs66.Fields(15).Value = True, "     Yes", "     No"), "S") & "," & GetStringToPrint(10, IIf(IsDBNull(chkrs66.Fields(13).Value), "     Due", "     Paid"), "S") & vbNewLine)

                                totamtt = totamtt + chkrs66.Fields(5).Value
                                totcgst = totcgst + chkrs66.Fields(7).Value
                                totsgst = totsgst + chkrs66.Fields(9).Value
                                totnet = totnet + chkrs66.Fields(10).Value
                                If chkrs66.Fields(4).Value <= ED Then
                                    advRec = advRec + chkrs66.Fields(10).Value
                                End If
                            End If
                            netInvoiceAmt = netInvoiceAmt + chkrs66.Fields(10).Value
                            If chkrs66.Fields(4).Value <= ED Then
                                advreceived = advreceived + chkrs66.Fields(10).Value
                            End If
                            If chkrs66.EOF = False Then
                                chkrs66.MoveNext()
                            End If
                        Loop

                        chkrs66.Close()
                        Print(fnum, StrDup(100, "=") & vbNewLine)
                        Print(fnumm, StrDup(100, "=") & vbNewLine)
                        Print(fnum, GetStringToPrint(11, "", "S") & GetStringToPrint(10, "", "S") & GetStringToPrint(13, totamtt, "N") & GetStringToPrint(7, " ", "S") & GetStringToPrint(12, totcgst, "N") & GetStringToPrint(12, totsgst, "N") & GetStringToPrint(12, totnet, "N") & GetStringToPrint(20, "", "S") & vbNewLine)
                        Print(fnumm, GetStringToPrint(11, "", "S") & "," & GetStringToPrint(10, "", "S") & "," & GetStringToPrint(13, totamtt, "N") & "," & GetStringToPrint(7, " ", "S") & "," & GetStringToPrint(12, totcgst, "N") & "," & GetStringToPrint(12, totsgst, "N") & "," & GetStringToPrint(12, totnet, "N") & "," & GetStringToPrint(20, "", "S") & vbNewLine)
                        '''''''''''''''''''' invoice details


                        ''''''''''''''''''''payment details
                        Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"
                        chkrs11.Open("SELECT DISTINCT [receipt].*,[bill].[P_CODE],[BILL].[REC_DATE],[BILL].[REC_NO] from [receipt] INNER JOIN [bill] on [receipt].[rec_no]=int([bill].[rec_no]) and [receipt].[rec_date]=[bill].[rec_date] WHERE [receipt].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [receipt].[GODWN_NO]='" & chkrs2.Fields(3).Value & "' and [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [receipt].REC_DATE <= FORMAT('" & DateTimePicker1.Value.ToString("dd/MM/yyyy") & "','DD/MM/YYYY') and [BILL].P_CODE='" & chkrs2.Fields(1).Value & "' order by [receipt].[GROUP],[receipt].GODWN_NO,[receipt].REC_DATE,[receipt].REC_NO", xcon)
                        If chkrs11.BOF = False Then
                            chkrs11.MoveFirst()
                        End If
                        Dim first As Boolean = True
                        Dim totamt As Double = 0
                        Dim totadv As Double = 0
                        Do While chkrs11.EOF = False
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' AND GODWN_NO='" & chkrs11.Fields(2).Value & "' AND [STATUS]='C'", xcon)

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
                            If chkrs44.EOF = False Then
                                If IsDBNull(chkrs44.Fields(5).Value) Then
                                Else
                                    CENSUS = chkrs44.Fields(5).Value
                                End If
                                If IsDBNull(chkrs44.Fields(4).Value) Then
                                Else
                                    SURVEY = chkrs44.Fields(4).Value
                                End If
                                pname = chkrs44.Fields(38).Value
                                pcode1 = chkrs44.Fields(1).Value

                                chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                Dim amtt As Double = 0
                                If chkrs22.EOF = False Then
                                    chkrs22.MoveFirst()
                                    amtt = chkrs22.Fields(4).Value
                                    If IsDBNull(chkrs22.Fields(5).Value) Then
                                    Else
                                        amtt = amtt + chkrs22.Fields(5).Value
                                    End If
                                End If
                                chkrs22.Close()
                                chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                If chkrs33.EOF = False Then
                                    If IsDBNull(chkrs33.Fields(2).Value) Then
                                        CGST_RATE = 0
                                    Else
                                        CGST_RATE = chkrs33.Fields(2).Value
                                    End If
                                    If IsDBNull(chkrs33.Fields(3).Value) Then
                                        SGST_RATE = 0
                                    Else
                                        SGST_RATE = chkrs33.Fields(3).Value
                                    End If
                                End If
                                gst = CGST_RATE + SGST_RATE
                                chkrs33.Close()
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
                            chkrs44.Close()

                            Dim grp As String = chkrs11.Fields(1).Value
                            Dim gdn As String = chkrs11.Fields(2).Value
                            Dim invdt As DateTime = chkrs11.Fields(3).Value
                            Dim inv As Integer = chkrs11.Fields(4).Value
                            Dim FIRSTREC As Boolean = True
                            Dim FROMNO As String = ""
                            Dim TONO As String
                            Dim against As String = ""
                            Dim against1 As String = ""
                            Dim against3 As String = ""
                            Dim against2 As String = ""
                            Dim agcount As Integer = 0
                            Dim adjusted_amt As Double = 0
                            Dim last_bldate As DateTime
                            Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                            chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                            Do While chkrs22.EOF = False
                                If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                    If FIRSTREC Then
                                        chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                        If chkrs6.EOF = False Then
                                            FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                            TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                        Else
                                            FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                            TONO = FROMNO
                                        End If
                                        chkrs6.Close()

                                        FIRSTREC = False
                                    Else
                                        TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                    End If
                                    pname = chkrs22.Fields(16).Value
                                    pcode1 = chkrs22.Fields(3).Value
                                    adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                    last_bldate = chkrs22.Fields(4).Value
                                    If agcount < 7 Then
                                        against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                    Else
                                        If agcount < 14 Then
                                            against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 21 Then
                                                against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 28 Then
                                                    against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                End If
                                            End If
                                        End If
                                    End If

                                    agcount = agcount + 1
                                End If
                                If chkrs22.EOF = False Then
                                    chkrs22.MoveNext()
                                End If
                                If chkrs22.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs22.Close()
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
                            advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                            advanceamtprint = advanceamt
                            If advanceamt > 0 Then
                                Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs55.Open(Rss, xcon)
                                Do While chkrs55.EOF = False
                                    If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                        lastbilladjusted = chkrs55.Fields(0).Value
                                        last_bldate = chkrs55.Fields(4).Value
                                    End If
                                    If chkrs55.EOF = False Then
                                        chkrs55.MoveNext()
                                    End If
                                Loop
                                chkrs55.Close()
                                If lastbilladjusted = 0 Then
                                    Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND P_CODE='" & pcode1 & "' and [STATUS]='C' order by [GROUP]+GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    If chkrs55.EOF = False Then
                                        last_bldate = chkrs55.Fields(0).Value
                                    End If
                                    chkrs55.Close()
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
                            recntAMT = recntAMT + chkrs11.Fields(5).Value
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If (chkrs11.Fields(3).Value >= Format(DateTimePicker5.Value.Date, "dd/MM/yyyy") And chkrs11.Fields(3).Value <= Format(DateTimePicker4.Value.Date, "dd/MM/yyyy")) Then
                                If first Then
                                    Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(50, "Tenant Name", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & vbNewLine)

                                    Print(fnum, GetStringToPrint(17, "Godown Type", "S") & GetStringToPrint(17, "Godown No.", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "  ", "S") & GetStringToPrint(12, " ", "S") & GetStringToPrint(50, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Godown Type", "S") & "," & GetStringToPrint(17, "Godown No.", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(12, " ", "S") & "," & GetStringToPrint(50, "Bank A/C Detail", "S") & vbNewLine)

                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, "Bank Name & Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, "Bank Name & Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(133, "=") & vbNewLine)
                                    Print(fnumm, StrDup(133, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                Else
                                    Print(fnum, StrDup(30, ">") & vbNewLine)
                                End If


                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(50, pname, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(50, pname, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(1).Value, "S") & GetStringToPrint(17, chkrs11.Fields(2).Value.ToString, "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(13, " ", "S") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(17, chkrs11.Fields(2).Value.ToString, "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.Fields(7).Value.Equals("C") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(50, chkrs11.Fields(8).Value & " " & chkrs11.Fields(9).Value, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, chkrs11.Fields(8).Value & " " & chkrs11.Fields(9).Value, "S") & vbNewLine)
                                    xline = xline + 1
                                End If




                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                xline = xline + 2

                            End If

                            If chkrs11.EOF = False Then
                                chkrs11.MoveNext()
                            End If
                            If chkrs11.EOF = True Then
                                Exit Do
                            End If

                        Loop
                        If totamt > 0 Then
                            Print(fnum, StrDup(133, "-") & vbNewLine)
                            Print(fnumm, StrDup(133, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(50, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(50, "", "S") & vbNewLine)
                        End If
                        chkrs11.Close()

                        ''''''''''''''''''''payment details
                    Else
                        ''''''''''''''''checking in godown transfer table
                        chkrs4.Open("SELECT [GDTRANS].*,[PARTY].P_NAME FROM GDTRANS INNER JOIN PARTY ON [PARTY].P_CODE=[GDTRANS].NP_CODE WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and OP_CODE ='" & TextBox1.Text & "' order by  [GROUP],[GODWN_NO],OP_CODE DESC", xcon)
                        If chkrs4.BOF = False Then
                            chkrs4.MoveFirst()
                        End If
                        Do While chkrs4.EOF = False
                            Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Transfered to ", "S") & GetStringToPrint(55, chkrs4.Fields(7).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Transfered to ", "S") & "," & GetStringToPrint(55, chkrs4.Fields(7).Value, "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(13, "Used From  : ", "S") & GetStringToPrint(12, chkrs2.Fields(11).Value, "S") & GetStringToPrint(9, "up to :", "S") & GetStringToPrint(15, chkrs4.Fields(6).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Used From  : ", "S") & "," & GetStringToPrint(12, chkrs2.Fields(11).Value, "S") & "," & GetStringToPrint(9, "up to :", "S") & "," & GetStringToPrint(15, chkrs4.Fields(6).Value, "S") & vbNewLine)
                            If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            End If
                            ''''''''''''''''''''payment details
                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Dim STR As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs4.Fields(6).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"
                            chkrs11.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs4.Fields(6).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
                            If chkrs11.BOF = False Then
                                chkrs11.MoveFirst()
                            End If
                            Dim first As Boolean = True
                            Dim totamt As Double = 0
                            Dim totadv As Double = 0
                            Do While chkrs11.EOF = False

                                If first Then

                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(180, "=") & vbNewLine)
                                    Print(fnumm, StrDup(180, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                End If
                                Dim srn As String = "SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE P_CODE='" & TextBox1.Text & "' AND [GROUP]='" & chkrs2.Fields(0).Value & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "'"

                                chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GODOWN].P_CODE='" & TextBox1.Text & "' AND [GROUP]='" & chkrs2.Fields(0).Value & "' AND GODWN_NO='" & chkrs2.Fields(3).Value & "'", xcon)

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
                                If chkrs44.EOF = False Then
                                    If IsDBNull(chkrs44.Fields(5).Value) Then
                                    Else
                                        CENSUS = chkrs44.Fields(5).Value
                                    End If
                                    If IsDBNull(chkrs44.Fields(4).Value) Then
                                    Else
                                        SURVEY = chkrs44.Fields(4).Value
                                    End If
                                    pname = chkrs44.Fields(38).Value
                                    pcode1 = TextBox1.Text

                                    chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                    Dim amtt As Double = 0
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveFirst()
                                        amtt = chkrs22.Fields(4).Value
                                        If IsDBNull(chkrs22.Fields(5).Value) Then
                                        Else
                                            amtt = amtt + chkrs22.Fields(5).Value
                                        End If
                                    End If
                                    chkrs22.Close()
                                    chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                    If chkrs33.EOF = False Then
                                        If IsDBNull(chkrs33.Fields(2).Value) Then
                                            CGST_RATE = 0
                                        Else
                                            CGST_RATE = chkrs33.Fields(2).Value
                                        End If
                                        If IsDBNull(chkrs33.Fields(3).Value) Then
                                            SGST_RATE = 0
                                        Else
                                            SGST_RATE = chkrs33.Fields(3).Value
                                        End If
                                    End If
                                    gst = CGST_RATE + SGST_RATE
                                    chkrs33.Close()
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
                                chkrs44.Close()

                                Dim grp As String = chkrs11.Fields(1).Value
                                Dim gdn As String = chkrs11.Fields(2).Value
                                Dim invdt As DateTime = chkrs11.Fields(3).Value
                                Dim inv As Integer = chkrs11.Fields(4).Value
                                Dim FIRSTREC As Boolean = True
                                Dim FROMNO As String = ""
                                Dim TONO As String
                                Dim against As String = ""
                                Dim against1 As String = ""
                                Dim against3 As String = ""
                                Dim against2 As String = ""
                                Dim agcount As Integer = 0
                                Dim adjusted_amt As Double = 0
                                Dim last_bldate As DateTime
                                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=FORMAT('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                                Do While chkrs22.EOF = False
                                    If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                        If FIRSTREC Then
                                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                            If chkrs6.EOF = False Then
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                            Else
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                                TONO = FROMNO
                                            End If
                                            chkrs6.Close()

                                            FIRSTREC = False
                                        Else
                                            TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                        End If
                                        pname = chkrs22.Fields(16).Value
                                        pcode1 = chkrs22.Fields(3).Value
                                        adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                        last_bldate = chkrs22.Fields(4).Value
                                        If agcount < 7 Then
                                            against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 14 Then
                                                against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 21 Then
                                                    against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                    If agcount < 28 Then
                                                        against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If

                                        agcount = agcount + 1
                                    End If
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveNext()
                                    End If
                                    If chkrs22.EOF = True Then
                                        Exit Do
                                    End If

                                Loop
                                chkrs22.Close()
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
                                advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                                advanceamtprint = advanceamt
                                If advanceamt > 0 Then
                                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    Do While chkrs55.EOF = False
                                        If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                            lastbilladjusted = chkrs55.Fields(0).Value
                                            last_bldate = chkrs55.Fields(4).Value
                                        End If
                                        If chkrs55.EOF = False Then
                                            chkrs55.MoveNext()
                                        End If
                                    Loop
                                    chkrs55.Close()
                                    If lastbilladjusted = 0 Then
                                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' and  P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                                        chkrs55.Open(Rss, xcon)
                                        If chkrs55.EOF = False Then
                                            last_bldate = chkrs55.Fields(0).Value
                                        End If
                                        chkrs55.Close()
                                    End If
                                    Dim dtcounter As Long = 1
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

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.EOF = False Then
                                    chkrs11.MoveNext()
                                End If
                                If chkrs11.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs11.Close()
                            Print(fnum, StrDup(180, "-") & vbNewLine)
                            Print(fnumm, StrDup(180, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                            ''''''''''''''''''''payment details
                            If chkrs4.BOF = False Then
                                chkrs4.MoveNext()
                            End If
                        Loop
                        chkrs4.Close()
                        '   chkrs6.Close()
                        ''''''''''''''''''''details of payment for godown


                        '''''''''''''''''''checking in godown close /suspend tabel
                        chkrs5.Open("SELECT * FROM CLGDWN WHERE [GROUP]='" & chkrs2.Fields(0).Value & "' and GODWN_NO='" & chkrs2.Fields(3).Value & "' and P_CODE ='" & TextBox1.Text & "' order by [GROUP],GODWN_NO,P_CODE DESC", xcon)
                        If chkrs5.BOF = False Then
                            chkrs5.MoveFirst()
                        End If
                        Do While chkrs5.EOF = False
                            If chkrs5.Fields(6).Value.Equals("S") Then
                                Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Suspended", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Suspended", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Status     : ", "S") & GetStringToPrint(15, "Closed", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Status     : ", "S") & "," & GetStringToPrint(15, "Closed", "S") & vbNewLine)
                            End If

                            Print(fnum, GetStringToPrint(13, "Used From  : ", "S") & GetStringToPrint(15, chkrs2.Fields(11).Value, "S") & GetStringToPrint(13, "up to  : ", "S") & GetStringToPrint(15, chkrs5.Fields(3).Value, "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(13, "Used From  : ", "S") & "," & GetStringToPrint(15, chkrs2.Fields(11).Value, "S") & "," & GetStringToPrint(13, "up to  : ", "S") & "," & GetStringToPrint(15, chkrs5.Fields(3).Value, "S") & vbNewLine)
                            If IsDBNull(chkrs5.Fields(5).Value) Or chkrs5.Fields(5).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Reason     : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Reason     : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Reason     : ", "S") & GetStringToPrint(55, chkrs5.Fields(5).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Reason     : ", "S") & "," & GetStringToPrint(55, chkrs5.Fields(5).Value, "S") & vbNewLine)
                            End If

                            If IsDBNull(chkrs2.Fields(22).Value) Or chkrs2.Fields(22).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, "Remarks    : ", "S") & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, "Remarks    : ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(22).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(23).Value) Or chkrs2.Fields(23).Value.Equals("") Then
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, " ", "S") & vbNewLine)
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(23).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(24).Value) Or chkrs2.Fields(24).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(24).Value, "S") & vbNewLine)
                            End If
                            If IsDBNull(chkrs2.Fields(25).Value) Or chkrs2.Fields(25).Value.Equals("") Then
                            Else
                                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(13, " ", "S") & "," & GetStringToPrint(55, chkrs2.Fields(25).Value, "S") & vbNewLine)
                            End If
                            ''''''''''''''''''''payment details
                            Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "Payment Details  ", "S") & vbNewLine)
                            Print(fnum, GetStringToPrint(30, "===============  ", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(30, "===============  ", "S") & vbNewLine)

                            chkrs11.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE >= FORMAT('" & chkrs2.Fields(11).Value & "','DD/MM/YYYY') AND [RECEIPT].REC_DATE < FORMAT('" & chkrs5.Fields(3).Value & "','DD/MM/YYYY') and [RECEIPT].[GROUP]='" & chkrs2.Fields(0).Value & "' AND [RECEIPT].GODWN_NO='" & chkrs2.Fields(3).Value & "' order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
                            If chkrs11.BOF = False Then
                                chkrs11.MoveFirst()
                            End If
                            Dim first As Boolean = True
                            Dim totamt As Double = 0
                            Dim totadv As Double = 0
                            Do While chkrs11.EOF = False

                                If first Then

                                    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Receipt Date", "S") & "," & GetStringToPrint(17, "Receipt No.", "S") & "," & GetStringToPrint(13, "Amount", "N") & "," & GetStringToPrint(13, "  Advance", "S") & "," & GetStringToPrint(12, "Cash/Cheque", "S") & "," & GetStringToPrint(7, "Group", "S") & "," & GetStringToPrint(13, "Godown No.", "S") & "," & GetStringToPrint(50, "Tenant Name", "S") & "," & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "Cheque No.", "S") & "," & GetStringToPrint(17, "Cheque Date", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, "Bank Name", "S") & "," & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                                    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "From Month-Year", "S") & "," & GetStringToPrint(17, "To Month-Year", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                                    Print(fnum, StrDup(180, "=") & vbNewLine)
                                    Print(fnumm, StrDup(180, "=") & vbNewLine)
                                    first = False
                                    xline = xline + 3
                                End If

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                chkrs44.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)

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
                                If chkrs44.EOF = False Then
                                    If IsDBNull(chkrs44.Fields(5).Value) Then
                                    Else
                                        CENSUS = chkrs44.Fields(5).Value
                                    End If
                                    If IsDBNull(chkrs44.Fields(4).Value) Then
                                    Else
                                        SURVEY = chkrs44.Fields(4).Value
                                    End If
                                    pname = chkrs44.Fields(38).Value
                                    pcode1 = TextBox1.Text

                                    chkrs22.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs11.Fields(1).Value & "' and GODWN_NO='" & chkrs11.Fields(2).Value & "' and P_CODE ='" & chkrs44.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                                    Dim amtt As Double = 0
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveFirst()
                                        amtt = chkrs22.Fields(4).Value
                                        If IsDBNull(chkrs22.Fields(5).Value) Then
                                        Else
                                            amtt = amtt + chkrs22.Fields(5).Value
                                        End If
                                    End If
                                    chkrs22.Close()
                                    chkrs33.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs44.Fields(37).Value & "'", xcon)

                                    If chkrs33.EOF = False Then
                                        If IsDBNull(chkrs33.Fields(2).Value) Then
                                            CGST_RATE = 0
                                        Else
                                            CGST_RATE = chkrs33.Fields(2).Value
                                        End If
                                        If IsDBNull(chkrs33.Fields(3).Value) Then
                                            SGST_RATE = 0
                                        Else
                                            SGST_RATE = chkrs33.Fields(3).Value
                                        End If
                                    End If
                                    gst = CGST_RATE + SGST_RATE
                                    chkrs33.Close()
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
                                chkrs44.Close()

                                Dim grp As String = chkrs11.Fields(1).Value
                                Dim gdn As String = chkrs11.Fields(2).Value
                                Dim invdt As DateTime = chkrs11.Fields(3).Value
                                Dim inv As Integer = chkrs11.Fields(4).Value
                                Dim FIRSTREC As Boolean = True
                                Dim FROMNO As String = ""
                                Dim TONO As String
                                Dim against As String = ""
                                Dim against1 As String = ""
                                Dim against3 As String = ""
                                Dim against2 As String = ""
                                Dim agcount As Integer = 0
                                Dim adjusted_amt As Double = 0
                                Dim last_bldate As DateTime
                                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                chkrs22.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                                Do While chkrs22.EOF = False
                                    If chkrs22.Fields(13).Value >= inv And chkrs22.Fields(14).Value <= invdt And chkrs11.Fields(3).Value >= chkrs22.Fields(4).Value Then
                                        If FIRSTREC Then
                                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs22.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs22.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                                            If chkrs6.EOF = False Then
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                                            Else
                                                FROMNO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                                TONO = FROMNO
                                            End If
                                            chkrs6.Close()

                                            FIRSTREC = False
                                        Else
                                            TONO = MonthName(Convert.ToDateTime(chkrs22.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs22.Fields(4).Value).Year
                                        End If
                                        pname = chkrs22.Fields(16).Value
                                        pcode1 = chkrs22.Fields(3).Value
                                        adjusted_amt = adjusted_amt + chkrs22.Fields(10).Value
                                        last_bldate = chkrs22.Fields(4).Value
                                        If agcount < 7 Then
                                            against = against + "GO-" & chkrs22.Fields(0).Value & ", "
                                        Else
                                            If agcount < 14 Then
                                                against1 = against1 + "GO-" & chkrs22.Fields(0).Value & ", "
                                            Else
                                                If agcount < 21 Then
                                                    against2 = against2 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                Else
                                                    If agcount < 28 Then
                                                        against3 = against3 + "GO-" & chkrs22.Fields(0).Value & ", "
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If

                                        agcount = agcount + 1
                                    End If
                                    If chkrs22.EOF = False Then
                                        chkrs22.MoveNext()
                                    End If
                                    If chkrs22.EOF = True Then
                                        Exit Do
                                    End If

                                Loop
                                chkrs22.Close()
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

                                advanceamt = chkrs11.Fields(5).Value - adjusted_amt
                                advanceamtprint = advanceamt
                                If advanceamt > 0 Then
                                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                                    chkrs55.Open(Rss, xcon)
                                    Do While chkrs55.EOF = False
                                        If chkrs11.Fields(3).Value >= chkrs55.Fields(4).Value Then
                                            lastbilladjusted = chkrs55.Fields(0).Value
                                            last_bldate = chkrs55.Fields(4).Value
                                        End If
                                        If chkrs55.EOF = False Then
                                            chkrs55.MoveNext()
                                        End If
                                    Loop
                                    chkrs55.Close()
                                    If lastbilladjusted = 0 Then
                                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND [STATUS]='C' AND P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                                        chkrs55.Open(Rss, xcon)
                                        If chkrs55.EOF = False Then
                                            last_bldate = chkrs55.Fields(0).Value
                                        End If
                                        chkrs55.Close()
                                    End If
                                    Dim dtcounter As Long = 1
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

                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                totamt = totamt + chkrs11.Fields(5).Value
                                totadv = totadv + advanceamtprint
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(3).Value, "S") & "," & GetStringToPrint(17, "GST-" + chkrs11.Fields(4).Value.ToString, "S") & "," & GetStringToPrint(13, Format(chkrs11.Fields(5).Value, "######0.00"), "N") & "," & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & "," & GetStringToPrint(12, "  " & chkrs11.Fields(7).Value, "S") & "," & GetStringToPrint(7, chkrs11.Fields(1).Value, "S") & "," & GetStringToPrint(13, chkrs11.Fields(2).Value, "S") & "," & GetStringToPrint(50, pname, "S") & "," & GetStringToPrint(33, IIf(IsDBNull(chkrs11.Fields(12).Value), "", chkrs11.Fields(12).Value), "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, chkrs11.Fields(10).Value, "S") & "," & GetStringToPrint(17, IIf(chkrs11.Fields(7).Value.Equals("C"), "", "  " & chkrs11.Fields(11).Value), "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(50, chkrs11.Fields(8).Value, "S") & "," & GetStringToPrint(33, chkrs11.Fields(9).Value, "S") & vbNewLine)
                                xline = xline + 1
                                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(17, FROMNO, "S") & "," & GetStringToPrint(17, TONO, "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against, "S") & vbNewLine)
                                xline = xline + 1
                                If Trim(against1).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against1, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against2).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against2, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                If Trim(against3).Equals("") Then
                                Else
                                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    Print(fnumm, GetStringToPrint(17, "", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, " ", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(12, "", "S") & "," & GetStringToPrint(7, "", "N") & "," & GetStringToPrint(13, "", "N") & "," & GetStringToPrint(63, against3, "S") & vbNewLine)
                                    xline = xline + 1
                                End If
                                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                                Print(fnumm, GetStringToPrint(15, " ", "S") & vbNewLine)
                                xline = xline + 1
                                If chkrs11.EOF = False Then
                                    chkrs11.MoveNext()
                                End If
                                If chkrs11.EOF = True Then
                                    Exit Do
                                End If

                            Loop
                            chkrs11.Close()
                            Print(fnum, StrDup(180, "-") & vbNewLine)
                            Print(fnumm, StrDup(180, "-") & vbNewLine)
                            Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
                            Print(fnumm, GetStringToPrint(17, "Total-->", "S") & "," & GetStringToPrint(17, "", "S") & "," & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & "," & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & "," & GetStringToPrint(12, "  ", "S") & "," & GetStringToPrint(7, "", "S") & "," & GetStringToPrint(13, "", "S") & "," & GetStringToPrint(50, "", "S") & "," & GetStringToPrint(33, "", "S") & vbNewLine)
                            ''''''''''''''''''''payment details
                            If chkrs5.BOF = False Then
                                chkrs5.MoveNext()
                            End If
                        Loop
                        chkrs5.Close()

                    End If


                End If
                Dim otst As Double = 0
                Dim prvdt As String = ""
                If IsDBNull(chkrs2.Fields(21).Value) Then
                Else
                    otst = chkrs2.Fields(21).Value
                End If

                prvdt = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToInt16(chkrs2.Fields(14).Value)).ToString() + " - " + chkrs2.Fields(15).Value.ToString()
                If chkrs2.BOF = False Then
                    chkrs2.MoveNext()
                End If
                Print(fnum, GetStringToPrint(20, "  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(20, "  ", "S") & vbNewLine)
                If (otst > 0) Then
                    Print(fnum, GetStringToPrint(42, "Previous Outstanding Detail as on Date : ", "S") & GetStringToPrint(15, Format(otst, "########0.00"), "N") & GetStringToPrint(30, " from " & prvdt, "S") & vbNewLine)
                    Print(fnumm, GetStringToPrint(42, "Previous Outstanding Detail as on Date : ", "S") & GetStringToPrint(15, Format(otst, "########0.00"), "N") & GetStringToPrint(30, " from " & prvdt, "S") & vbNewLine)
                End If
                Print(fnum, GetStringToPrint(42, "Outstanding Detail as on Date : " + DateTimePicker1.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(42, "Outstanding Detail as on Date : " + DateTimePicker1.Value.ToShortDateString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(42, "===============================================  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(42, "===============================================  ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Advance Opening -->  " & advreceived, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Advance Opening -->  " & advreceived, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Receipt tot     -->  " & recntAMT, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Receipt tot     -->  " & recntAMT, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "Invoice Amt     -->  " & netInvoiceAmt, "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "Invoice Amt     -->  " & netInvoiceAmt, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(30, "===================  ", "S") & vbNewLine)
                Print(fnumm, GetStringToPrint(30, "===================  ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(21, "Outstanding Amt -->  ", "S"))
                Print(fnumm, GetStringToPrint(21, "Outstanding Amt -->  ", "S"))
                Print(fnum, GetStringToPrint(30, IIf((netInvoiceAmt - recntAMT - advreceived) > 0, (netInvoiceAmt - recntAMT - advreceived), ((netInvoiceAmt - recntAMT - advreceived) * -1) & " Advance"), "S"))
                Print(fnumm, GetStringToPrint(30, IIf((netInvoiceAmt - recntAMT - advreceived) > 0, (netInvoiceAmt - recntAMT - advreceived), ((netInvoiceAmt - recntAMT - advreceived) * -1) & " Advance"), "S"))
            Loop
            chkrs2.Close()
            chkrs1.Close()

            FileClose(fnum)
            FileClose(fnumm)
            MyConn.Close()
            FrmGdnDtlView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Godowndetail.dat", RichTextBoxStreamType.PlainText)

            FrmGdnDtlView.Show()
            CreatePDF(Application.StartupPath & "\Reports\Godowndetail.dat", Application.StartupPath & "\Reports\Godowndetail")

            Dim PrintPDFFile As New ProcessStartInfo
            PrintPDFFile.UseShellExecute = True

            PrintPDFFile.Verb = "print"
            PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
            MsgBox(Application.StartupPath + " \Reports\" & TextBox5.Text & ".csv file is generated")

            '.Hidden
            PrintPDFFile.FileName = Application.StartupPath & "\Reports\Godowndetail.pdf"
            Process.Start(PrintPDFFile)
            '''''''''''''''''
        End If
    End Sub

    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        '''''''''search godown datagrid for the text user type in search text box
        MyConn = New OleDbConnection(connString)
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [godown].[status]='C' and " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")
        DataGridView2.DataSource = ds.Tables("GODOWN")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
    End Sub

End Class