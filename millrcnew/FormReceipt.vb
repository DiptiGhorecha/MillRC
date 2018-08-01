Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Public Class FormReceipt
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
    Public rownum As Integer = 0
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim dagp As OleDbDataAdapter
    Dim dsgp As DataSet
    Dim indexorder As String = "GODWN_NO"
    Private frmload As Boolean = True
    Dim GrpAddCorrect As String
    Private bValidateinvoice As Boolean = True
    Private bValidatetype As Boolean = True
    Private bValidategodown As Boolean = True
    Private bValidateamount As Boolean = True
    Private bValidatedate As Boolean = True
    Dim formloaded As Boolean = False
    Dim checkinserted As Boolean = False
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim save As String
    Dim pdfpath As String
    Dim strReportFilePath As String
    Public FILE_NO As String
    Dim ctrlname As String = "ComboBox4"
    Dim payable As Double = 0
    Dim groupfilled As Boolean = False
    Dim godownfilled As Boolean = False
    Public lastdate As Date = DateTime.Today
    Public gridline As Integer
    Public rentsuggestion As Boolean = False

    Private Sub FormReceipt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        cmdAdd.Enabled = True
        cmdClose.Enabled = True
        cmdDelete.Enabled = True
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = False
        cmdCancel.Enabled = False
        DateTimePicker1.MaxDate = Date.Today
        disablefields()
        fillgroupcombo()
        groupfilled = True
        fillgodowncombo()
        godownfilled = True
        fillgroupbox2()
        ShowData()
        If DataGridView1.RowCount >= 1 Then
            rownum = DataGridView1.RowCount - 1
        End If
        LodaDataToTextBox()
        ' GroupBox1.Visible = True
        '  cmdFirst_Click(Nothing, Nothing)

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
    Private Sub FormReceipt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 And ((Me.ActiveControl.Name = "ComboBox4") Or Me.ActiveControl.Name = "ComboBox3") Then
            'GroupBox2.BringToFront()
            'GroupBox2.Visible = True
            'Label17.Text = "Godown Detail"
            'ctrlname = Me.ActiveControl.Name
            'GroupBox2.Focus()
            helpgrpcombo = ComboBox3
            helpgdncombo = ComboBox4
            GodownHelp.Show()

            '  ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        End If
        If e.KeyCode = Keys.F5 And (Me.ActiveControl.Name = "TextBox1") Then

            frmcalculator.Show()

        End If
    End Sub
    Function fillgroupbox2()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [STATUS]='C' AND [GROUP]='" & ComboBox3.Text & "' order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView3.DataSource = ds.Tables(0).DefaultView
            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            DataGridView3.Columns(1).Visible = False
            DataGridView3.Columns(2).Visible = False
            DataGridView3.Columns(4).Visible = False
            DataGridView3.Columns(5).Visible = False
            DataGridView3.Columns(6).Visible = False
            DataGridView3.Columns(7).Visible = False
            DataGridView3.Columns(8).Visible = False
            DataGridView3.Columns(9).Visible = False
            DataGridView3.Columns(10).Visible = False
            DataGridView3.Columns(11).Visible = False
            DataGridView3.Columns(12).Visible = False
            DataGridView3.Columns(13).Visible = False
            DataGridView3.Columns(14).Visible = False
            DataGridView3.Columns(15).Visible = False
            DataGridView3.Columns(16).Visible = False
            DataGridView3.Columns(17).Visible = False
            DataGridView3.Columns(18).Visible = False
            DataGridView3.Columns(19).Visible = False
            DataGridView3.Columns(20).Visible = False
            DataGridView3.Columns(21).Visible = False
            DataGridView3.Columns(22).Visible = False
            DataGridView3.Columns(23).Visible = False
            DataGridView3.Columns(24).Visible = False
            DataGridView3.Columns(25).Visible = False
            DataGridView3.Columns(26).Visible = False
            DataGridView3.Columns(27).Visible = False
            DataGridView3.Columns(28).Visible = False
            DataGridView3.Columns(29).Visible = False
            DataGridView3.Columns(30).Visible = False
            DataGridView3.Columns(31).Visible = False
            DataGridView3.Columns(32).Visible = False
            DataGridView3.Columns(33).Visible = False
            DataGridView3.Columns(34).Visible = False
            DataGridView3.Columns(35).Visible = False
            DataGridView3.Columns(36).Visible = False
            DataGridView3.Columns(37).Visible = False
            DataGridView3.Columns(0).Visible = True
            DataGridView3.Columns(3).Visible = True
            DataGridView3.Columns(38).Visible = True
            DataGridView3.Columns(0).HeaderText = "Group"
            DataGridView3.Columns(0).Width = 51
            DataGridView3.Columns(3).Width = 71
            DataGridView3.Columns(38).Width = 405
            DataGridView3.Columns(3).HeaderText = "Godown"
            DataGridView3.Columns(38).HeaderText = "Tenant"
            DataGridView3.Columns(21).HeaderText = "Outstanding"
            DataGridView3.Columns(21).Width = 105
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub DataGridView3_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView3.DoubleClick
        ComboBox4.Text = GetValue(DataGridView3.Item(3, DataGridView3.CurrentRow.Index).Value)
        ComboBox4.SelectedIndex = ComboBox4.FindStringExact(ComboBox4.Text)
        Label14.Text = DataGridView3.Item(38, DataGridView3.CurrentRow.Index).Value
        If DataGridView3.Item(37, DataGridView3.CurrentRow.Index).Value.Equals("997212") Then
            Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
        Else
            Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
        End If
        GroupBox1.BringToFront()
        ' GroupBox1.Visible = True
        GroupBox2.SendToBack()
        Label17.Text = "Receipt Detail"
        '  GroupBox1.Visible = False
    End Sub
    Private Sub TextBox7_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TextBox7.Text & "%' and [STATUS]='C' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")
        DataGridView3.DataSource = ds.Tables("GODOWN")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
    End Sub
    Private Sub DataGridView3_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then
            indexorder = "[GODOWN].GROUP"
            GroupBox6.Text = "Search by Group Type"
            '    DataGridView2.Sort(DataGridView2.Columns(0), SortOrder.Descending)
        End If
        If e.ColumnIndex = 3 Then
            indexorder = "[GODOWN].GODWN_NO"
            GroupBox6.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 38 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox6.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
    End Sub
    Function disablefields()
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Enabled = False
        CheckBox1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        DataGridView2.Enabled = False
        RentComboBox.Enabled = False
        '  GroupBox1.Visible = True
        Label12.Text = "Bank Account Detail"
    End Function
    Function textenable()
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        CheckBox1.Enabled = False
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        DataGridView2.Enabled = True
        RentComboBox.Enabled = True
        '  GroupBox1.Visible = True
        Label12.Text = "Bank Account Detail"
    End Function
    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        rownum = DataGridView1.CurrentRow.Index
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        rownum = DataGridView1.CurrentRow.Index
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        rownum = DataGridView1.CurrentRow.Index
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub LodaDataToTextBox()
        Try
            Dim i As Integer
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            DateTimePicker1.Text = ""
            DateTimePicker2.Text = ""
            If DataGridView1.RowCount >= 1 Then
                DataGridView1.ClearSelection()
                DataGridView1.Rows(rownum).Selected = True
                DataGridView1.FirstDisplayedScrollingRowIndex = rownum
                DataGridView1.CurrentCell = DataGridView1.Rows(rownum).Cells(1)

                'If frmload = True Then
                '    i = 0
                '    frmload = False
                'Else
                '    i = DataGridView1.CurrentRow.Index
                'End If
                If rownum > 0 Then
                    i = rownum
                Else
                    i = DataGridView1.CurrentRow.Index

                End If

                '  i = DataGridView1.CurrentRow.Index

                '    fillgrid2(DataGridView1.Item(1, i).Value, DataGridView1.Item(2, i).Value, DataGridView1.Item(4, i).Value, DataGridView1.Item(3, i).Value.ToString)
                If Not IsDBNull(DataGridView1.Item(1, i).Value) Then
                    ComboBox3.Text = GetValue(DataGridView1.Item(1, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(2, i).Value) Then
                    '    TextBox1.Text = GetValue(DataGridView1.Item(2, i).Value)
                    ComboBox4.Text = GetValue(DataGridView1.Item(2, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(3, i).Value) Then
                    DateTimePicker1.Value = GetValue(DataGridView1.Item(3, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(4, i).Value) Then
                    TextBox2.Text = GetValue(DataGridView1.Item(4, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(5, i).Value) Then

                    TextBox1.Text = Format(CSng(GetValue(DataGridView1.Item(5, i).Value)), "#####0.00")

                    End If
                    If Not IsDBNull(DataGridView1.Item(6, i).Value) Then
                    CheckBox1.Checked = Convert.ToBoolean(DataGridView1.Item(6, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(7, i).Value) Then
                    If DataGridView1.Item(7, i).Value.Equals("Q") Then
                        RadioButton1.Checked = True
                        Label12.Text = "Bank Account Detail"
                    Else
                        RadioButton2.Checked = True
                        Label12.Text = "Paid By (with Phone Number)"
                    End If

                End If
                If Not IsDBNull(DataGridView1.Item(8, i).Value) Then
                    TextBox4.Text = GetValue(DataGridView1.Item(8, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(9, i).Value) Then
                    TextBox5.Text = GetValue(DataGridView1.Item(9, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(10, i).Value) Then
                    TextBox3.Text = GetValue(DataGridView1.Item(10, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(12, i).Value) Then
                    TextBox6.Text = GetValue(DataGridView1.Item(12, i).Value)
                End If
                If Not IsDBNull(DataGridView1.Item(11, i).Value) Then
                    DateTimePicker2.Value = GetValue(DataGridView1.Item(11, i).Value)
                End If

                If Not IsDBNull(DataGridView1.Item(1, i).Value) Then
                    If xcon.State = ConnectionState.Open Then
                    Else
                        xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                    End If
                    chkrs4.Open("sELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & GetValue(DataGridView1.Item(1, i).Value) & "' AND GODWN_NO='" & GetValue(DataGridView1.Item(2, i).Value) & "' AND  [GODOWN].[P_CODE]='" & GetValue(DataGridView1.Item(14, i).Value) & "'", xcon)
                    If chkrs4.EOF = False Then
                        Label14.Text = chkrs4.Fields(38).Value
                        If chkrs4.Fields(37).Value.Equals("997212") Then
                            Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                        Else
                            Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                        End If
                    End If
                    chkrs4.Close()
                    xcon.Close()
                    fillgrid2(DataGridView1.Item(1, i).Value, DataGridView1.Item(2, i).Value, DataGridView1.Item(4, i).Value, DataGridView1.Item(3, i).Value.ToString)
                End If
            End If
            Label21.Text = "Total : " & DataGridView1.RowCount  '- 1
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub
    Function fillgrid2(grp As String, gdn As String, inv As Integer, invdt As String)
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        '    Dim STR As String = "SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND REC_NO <> 0 AND REC_DATE is not NULL order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO"
        'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND REC_NO IS NULL AND REC_DATE is NULL order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
        da = New OleDb.OleDbDataAdapter("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(
 SELECT SUM(NET_AMOUNT) 
 FROM [BILL] as t1 
 WHERE t1.[GROUP]='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND ((t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy')) or (t1.REC_NO IS NULL AND t1.REC_DATE is NULL)) AND t1.BILL_DATE <=t2.BILL_DATE 
) AS balance,IIF(t2.rec_no is not null,TRUE,FALSE) AS checker from [BILL] AS t2 INNER JOIN [PARTY] on t2.P_CODE=[PARTY].P_CODE where t2.[GROUP]='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy')) or (t2.REC_NO IS NULL AND t2.REC_DATE is NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", MyConn)
        Dim str As String = "SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT)FROM [BILL] as t1 WHERE t1.[GROUP]='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND ((t1.REC_NO='" & inv & "' and  t1.REC_DATE=format(" & Convert.ToDateTime(invdt) & ",'dd/mm/yyyy')) or (t1.REC_NO IS NULL AND t1.REC_DATE is NULL)) AND t1.BILL_DATE <=t2.BILL_DATE) AS balance,IIF(t2.rec_no is not null,TRUE,FALSE) AS checker from [BILL] AS t2 INNER JOIN [PARTY] on t2.P_CODE=[PARTY].P_CODE where t2.[GROUP]='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#) or (t2.REC_NO IS NULL AND t2.REC_DATE is NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close()
        If DataGridView2.Columns.Contains("chk") Then
            DataGridView2.Columns.Remove("chk")
        End If
        If DataGridView2.RowCount >= 1 Then
            DataGridView2.Columns(0).Visible = False
            DataGridView2.Columns(2).Visible = False
            DataGridView2.Columns(4).Visible = True
            DataGridView2.Columns(15).Visible = False
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(16).HeaderText = "Cumulative Amt"
            '  DataGridView2.Columns(1).HeaderText = "Group"
            DataGridView2.Columns(16).Width = 130
            DataGridView2.Columns(1).Width = 51
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(15).Width = 300
            '   DataGridView2.Columns(2).HeaderText = "Godown"
            DataGridView2.Columns(10).HeaderText = "Amt"
            DataGridView2.Columns(4).HeaderText = "Invoice Date"
            DataGridView2.Columns(3).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(7).Visible = False
            DataGridView2.Columns(8).Visible = False
            DataGridView2.Columns(9).Visible = False
            DataGridView2.Columns(10).Visible = True
            DataGridView2.Columns(10).DefaultCellStyle.Format = "N2"
            DataGridView2.Columns(16).DefaultCellStyle.Format = "N2"
            DataGridView2.Columns(11).Visible = False
            DataGridView2.Columns(12).Visible = False
            DataGridView2.Columns(13).Visible = False
            DataGridView2.Columns(14).Visible = False
            DataGridView2.Columns(4).Width = 80
            DataGridView2.ReadOnly = False
            ' DataGridView2.Columns(0).ReadOnly = True
            DataGridView2.Columns(1).ReadOnly = True
            DataGridView2.Columns(10).ReadOnly = True
            DataGridView2.Columns(4).ReadOnly = True
            DataGridView2.Columns(15).ReadOnly = True
            DataGridView2.Columns(16).ReadOnly = True
            DataGridView2.Columns(0).Width = 60
            Dim chk As New DataGridViewCheckBoxColumn()
            ' DataGridView2.Columns.Add(chk)
            chk.HeaderText = "Select Bill"
            chk.Name = "chk"
            chk.ValueType = GetType(Boolean)
            chk.DataPropertyName = "checker"
            DataGridView2.Columns.Insert(0, chk)
            DataGridView2.Columns(0).Width = 60

            DataGridView2.Columns(0).ReadOnly = False

            For X As Integer = 0 To DataGridView2.RowCount - 1
                If IsDBNull(DataGridView2.Item(14, X).Value) Then
                    DataGridView2.Item(0, X).Value = False
                Else
                    DataGridView2.CurrentCell = DataGridView2.Item(0, X)
                    DataGridView2.BeginEdit(False)
                    DataGridView2.Item(0, X).Value = True
                    ' Dim DGVevent As New Windows.Forms.DataGridViewCellEventArgs(0, X)
                    ' DataGridView2_CellContentClick(DataGridView2, DGVevent)
                End If
            Next
            checkinserted = True
            RentComboBox.DisplayMember = "Rent"
            RentComboBox.ValueMember = "Month"
            Dim tb As New DataTable
            tb.Columns.Add("Month", GetType(Integer))
            tb.Columns.Add("Rent", GetType(String))
            For i As Integer = 1 To 24
                tb.Rows.Add(i, Convert.ToDateTime(DataGridView2.Item(5, (DataGridView2.RowCount - 1)).Value).AddMonths(i) & " - " & (Convert.ToDouble(DataGridView2.Item(17, DataGridView2.RowCount - 1).Value) + (Convert.ToDouble(DataGridView2.Item(11, 0).Value) * i)).ToString)
            Next
            RentComboBox.DataSource = tb

        Else
            '''''''''''''''''all bills paid than check net_amount of previous bill and store as paayble

            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            chkrs.Open("SELECT INVOICE_NO,GROUP,GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO,REC_NO,REC_DATE from [BILL] where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", xcon)
            Do While chkrs.EOF = False
                payable = chkrs.Fields(10).Value
                Exit Do
            Loop
            chkrs.Close()
            xcon.Close()
        End If

    End Function
    Private Sub cmdFirst_Click(sender As Object, e As EventArgs) Handles cmdFirst.Click
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(0).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
        rownum = 0
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow > 0 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow - 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow - 1).Cells(1)
            rownum = intRow - 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow < DataGridView1.RowCount - 1 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow + 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow + 1).Cells(1)
            rownum = intRow + 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(DataGridView1.RowCount - 1).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(1)
        rownum = DataGridView1.RowCount - 1
        LodaDataToTextBox()
    End Sub
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    Private Sub navigatedisable()
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
        TxtSrch.Enabled = False
    End Sub

    Private Sub navigateenable()
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        TxtSrch.Enabled = True
    End Sub
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            ' Dim STR As String = "SELECT [RECEIPT].*,[S].[P_CODE]  from [RECEIPT] INNER JOIN (SELECT [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE] FROM [BILL] GROUP BY [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE]) AS S ON [RECEIPT].[REC_DATE]=[S].[REC_DATE] AND TRIM(STR([RECEIPT].[REC_NO]))=[S].[REC_NO] order by [RECEIPT].[REC_DATE],[RECEIPT].REC_NO"
            Dim STR As String = "SELECT [RECEIPT].*,[S].[P_CODE] from [RECEIPT] LEFT OUTER JOIN (SELECT [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE] FROM [BILL] GROUP BY [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE]) AS S ON [RECEIPT].[REC_DATE]=[S].[REC_DATE] AND TRIM(STR([RECEIPT].[REC_NO]))=[S].[REC_NO] order by [RECEIPT].[REC_DATE],[RECEIPT].REC_NO"
            If TxtSrch.Text.Trim <> "" Then
                TxtSrch.Text = ""
                'STR = "SELECT [RECEIPT].*,[S].[P_CODE] from [RECEIPT] LEFT OUTER JOIN (SELECT [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE] FROM [BILL] GROUP BY [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE]) AS S ON [RECEIPT].[REC_DATE]=[S].[REC_DATE] AND TRIM(STR([RECEIPT].[REC_NO]))=[S].[REC_NO] where " & indexorder & " Like '%" & TxtSrch.Text & "%'  order by [RECEIPT].[REC_DATE],[RECEIPT].REC_NO"
            End If

            da = New OleDb.OleDbDataAdapter(STR, MyConn)
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
    End Sub
    Public Function fillgodowncombo()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dagp = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] WHERE [GROUP]='" & ComboBox3.Text & "' and [STATUS]='C' Order by GODWN_NO", MyConn)
            dsgp = New DataSet
            dsgp.Clear()
            dagp.Fill(dsgp, "GODOWN")
            ComboBox4.DataSource = dsgp.Tables("GODOWN")
            ComboBox4.DisplayMember = "GODWN_NO"
            ComboBox4.ValueMember = "GODWN_NO"
            dagp.Dispose()
            dsgp.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Godown combo fill :" & ex.Message)
        End Try
    End Function
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
            ComboBox3.DataSource = dsg.Tables("GROUP")
            ComboBox3.DisplayMember = "G_CODE"
            ComboBox3.ValueMember = "G_CODE"
            dag.Dispose()
            dsg.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Group combo fill :" & ex.Message)
        End Try
    End Function

    Private Sub FormReceipt_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            Label22.Text = "Add"
            DataGridView2.Enabled = True
            DataGridView1.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            TextBox2.Text = ""
            TextBox1.Enabled = True
            TextBox1.Text = "0.00"
            TextBox2.Enabled = True
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            ComboBox4.Enabled = True
            ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
            ComboBox4.Text = ""
            ComboBox3.Enabled = True
            '  ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox3.Text = ""
            Label14.Text = ""
            Label15.Text = ""
            Label19.Text = ""
            Label20.Text = ""

            ComboBox3.Select()
            CheckBox1.Checked = False
            RadioButton1.Checked = True
            ' DateTime.Today

            If IsDBNull(lastdate) Then
                DateTimePicker1.Value = DateTime.Today
                lastdate = DateTimePicker1.Value
            Else
                DateTimePicker1.Value = lastdate
            End If
            DateTimePicker2.Value = DateTime.Today
            TextBox2.Text = getinvoicesr()
            DataGridView2.DataSource = Nothing
            If DataGridView2.Columns.Contains("chk") Then
                DataGridView2.Columns.Remove("chk")
            End If
            DataGridView2.Refresh()
            ' DataGridView2.Rows.Clear()
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If groupfilled Then
            fillgodowncombo()
            fillgroupbox2()
        End If
    End Sub

    Public Function getinvoicesr()
        Try
            Dim INVNO As String
            Dim INVNOTMP As String
            Dim nom As Integer

            Dim iDate As String
            Dim fDate As DateTime
            Dim oDate As String
            Dim foDate As DateTime
            If (Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) >= 4) Then
                iDate = "01/04/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) + 1))
                foDate = Convert.ToDateTime(oDate)
            Else
                iDate = "01/04/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) - 1))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                foDate = Convert.ToDateTime(oDate)
            End If

            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            '  Dim SRD As String = "SELECT [RECEIPT].* from [RECEIPT] where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND Year([RECEIPT].REC_DATE)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [RECEIPT].REC_NO"
            'chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where Year([RECEIPT].REC_DATE)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [RECEIPT].REC_NO", xcon)
            chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE>=Format('" & fDate & "', 'dd/mm/yyyy') and [RECEIPT].REC_DATE<=Format('" & foDate & "', 'dd/mm/yyyy') order by [RECEIPT].REC_NO", xcon)
            ' chkrs1.MoveLast()
            Do While chkrs1.EOF = False
                INVNO = chkrs1.Fields(4).Value.ToString
                INVNOTMP = chkrs1.Fields(4).Value.ToString
                chkrs1.MoveNext()
            Loop
            chkrs1.Close()
            xcon.Close()
            nom = Convert.ToInt32(INVNO) + 1
            INVNO = nom     'String.Format("{0:000}", nom)
            Return INVNO
        Catch ex As Exception
            MsgBox("Exception: Get invoice SR :" & ex.Message)
        End Try

    End Function

    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
        Try
            If godownfilled Then
                If ComboBox4.Text.Equals("") Then
                Else
                    rentsuggestion = False
                    MyConn = New OleDbConnection(connString)
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    '    Dim STR As String = "SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND REC_NO <> 0 AND REC_DATE is not NULL order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO"
                    'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND REC_NO IS NULL AND REC_DATE is NULL order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
                    da = New OleDb.OleDbDataAdapter("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(
 SELECT SUM(NET_AMOUNT) 
 FROM [BILL] as t1 
 WHERE t1.[GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND t1.GODWN_NO='" & ComboBox4.SelectedValue.ToString & "' AND t1.REC_NO IS NULL AND t1.REC_DATE is NULL AND t1.BILL_DATE <=t2.BILL_DATE 
) AS balance,IIF(t2.rec_no is not null,TRUE,FALSE) AS checker from [BILL] AS t2 INNER JOIN [PARTY] on t2.P_CODE=[PARTY].P_CODE where t2.[GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND t2.GODWN_NO='" & ComboBox4.SelectedValue.ToString & "' AND t2.REC_NO IS NULL AND t2.REC_DATE is NULL order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", MyConn)
                    ds = New DataSet
                    ds.Clear()
                    da.Fill(ds)
                    DataGridView2.DataSource = ds.Tables(0).DefaultView
                    da.Dispose()
                    ds.Dispose()
                    MyConn.Close()
                    If DataGridView2.Columns.Contains("chk") Then
                        DataGridView2.Columns.Remove("chk")
                    End If
                    If DataGridView2.RowCount >= 1 Then
                        DataGridView2.Columns(0).Visible = False
                        DataGridView2.Columns(2).Visible = False
                        DataGridView2.Columns(4).Visible = True
                        DataGridView2.Columns(15).Visible = False
                        DataGridView2.Columns(1).Visible = False
                        DataGridView2.Columns(16).HeaderText = "Cumulative Amt"
                        '  DataGridView2.Columns(1).HeaderText = "Group"
                        DataGridView2.Columns(16).Width = 130
                        DataGridView2.Columns(1).Width = 51
                        DataGridView2.Columns(2).Width = 71
                        DataGridView2.Columns(15).Width = 300
                        '   DataGridView2.Columns(2).HeaderText = "Godown"
                        DataGridView2.Columns(10).HeaderText = "Amt"
                        DataGridView2.Columns(4).HeaderText = "Invoice Date"
                        DataGridView2.Columns(3).Visible = False
                        DataGridView2.Columns(5).Visible = False
                        DataGridView2.Columns(6).Visible = False
                        DataGridView2.Columns(7).Visible = False
                        DataGridView2.Columns(8).Visible = False
                        DataGridView2.Columns(9).Visible = False
                        DataGridView2.Columns(10).Visible = True
                        DataGridView2.Columns(10).DefaultCellStyle.Format = "N2"
                        DataGridView2.Columns(16).DefaultCellStyle.Format = "N2"
                        DataGridView2.Columns(11).Visible = False
                        DataGridView2.Columns(12).Visible = False
                        DataGridView2.Columns(13).Visible = False
                        DataGridView2.Columns(14).Visible = False
                        DataGridView2.Columns(4).Width = 80
                        DataGridView2.ReadOnly = False
                        DataGridView2.Columns(0).ReadOnly = True
                        DataGridView2.Columns(1).ReadOnly = True
                        DataGridView2.Columns(10).ReadOnly = True
                        DataGridView2.Columns(4).ReadOnly = True
                        DataGridView2.Columns(15).ReadOnly = True
                        DataGridView2.Columns(16).ReadOnly = True
                        '   DataGridView2.Columns(0).Width = 60
                        Dim chk As New DataGridViewCheckBoxColumn()
                        ' DataGridView2.Columns.Add(chk)
                        chk.HeaderText = "Select Bill"
                        chk.Name = "chk"
                        chk.DataPropertyName = "checker"
                        DataGridView2.Columns.Insert(0, chk)
                        DataGridView2.Columns(0).Width = 60
                        DataGridView2.Columns(0).ReadOnly = False
                        For X As Integer = 0 To DataGridView2.RowCount - 1
                            If IsDBNull(DataGridView2.Item(14, X).Value) Then
                                DataGridView2.Item(0, X).Value = False
                            Else
                                DataGridView2.CurrentCell = DataGridView2.Item(0, X)
                                DataGridView2.BeginEdit(False)
                                DataGridView2.Item(0, X).Value = True
                            End If
                        Next
                        checkinserted = True
                        If DataGridView2.CurrentRow Is Nothing Then
                            Label14.Text = DataGridView2.Item(16, 0).Value
                            If DataGridView2.Item(12, 0).Value.Equals("997212") Then
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                            Else
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                            End If
                        Else
                            Label14.Text = DataGridView2.Item(16, DataGridView2.CurrentRow.Index).Value
                            '   TextBox6.Text = DataGridView2.Item(16, DataGridView2.CurrentRow.Index).Value
                            If DataGridView2.Item(12, DataGridView2.CurrentRow.Index).Value.Equals("997212") Then
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                            Else
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                            End If
                        End If
                        Label19.Text = Format(Convert.ToDouble(DataGridView2.Item(11, 0).Value), "0.00")
                        Label20.Text = DataGridView2.Item(6, 0).Value.ToString & " + " & (Convert.ToDouble(DataGridView2.Item(8, 0).Value) + Convert.ToDouble(DataGridView2.Item(10, 0).Value)).ToString
                        ' MsgBox(DataGridView2.Item(17, DataGridView2.RowCount - 1).Value)
                        ' MsgBox(DataGridView2.Item(5, DataGridView2.RowCount - 1).Value)
                        RentComboBox.DisplayMember = "Rent"
                        RentComboBox.ValueMember = "Month"
                        Dim tb As New DataTable
                        tb.Columns.Add("Month", GetType(Integer))
                        tb.Columns.Add("Rent", GetType(String))
                        For i As Integer = 1 To 24
                            tb.Rows.Add(i, Convert.ToDateTime(DataGridView2.Item(5, (DataGridView2.RowCount - 1)).Value).AddMonths(i) & " - " & (Convert.ToDouble(DataGridView2.Item(17, DataGridView2.RowCount - 1).Value) + (Convert.ToDouble(DataGridView2.Item(11, 0).Value) * i)).ToString)
                        Next
                        RentComboBox.DataSource = tb
                    Else
                        '''''''''''''''''all bills paid than check net_amount of previous bill and store as paayble

                        If xcon.State = ConnectionState.Open Then
                        Else
                            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                        End If
                        Dim PCD As String
                        Dim hsnm As String
                        chkrs.Open("SELECT [GODOWN].*,[PARTY].P_NAME,[PARTY].GST from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].[GROUP]='" & ComboBox3.Text & "' AND [GODOWN].GODWN_NO='" & ComboBox4.Text & "' AND [GODOWN].STATUS='C' order by [GODOWN].GROUP,[GODOWN].GODWN_NO", xcon)
                        ' Select Case [BILL].*,[PARTY].P_NAME,[PARTY].GST from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO
                        Do While chkrs.EOF = False
                            ' payable = chkrs.Fields(10).Value
                            PCD = chkrs.Fields(1).Value
                            Label14.Text = Replace(chkrs.Fields(38).Value, "&", "&&")
                            hsnm = chkrs.Fields(37).Value
                            If hsnm.Equals("997211") Then
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                            Else
                                Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                            End If
                            Exit Do
                        Loop
                        chkrs.Close()
                        chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox3.Text & "' and GODWN_NO='" & ComboBox4.Text & "' and P_CODE ='" & PCD & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                        Dim amt As Double = 0
                        If chkrs2.EOF = False Then
                            chkrs2.MoveFirst()
                            amt = chkrs2.Fields(4).Value
                            If IsDBNull(chkrs2.Fields(5).Value) Then
                            Else
                                amt = amt + chkrs2.Fields(5).Value
                            End If
                        End If
                        chkrs2.Close()
                        payable = amt

                        chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & hsnm & "'", xcon)
                        Dim CGST_RATE As Double = 0
                        Dim SGST_RATE As Double = 0
                        Dim CGST_TAXAMT As Double = 0
                        Dim SGST_TAXAMT As Double = 0
                        Dim GST As Double = 0
                        Dim GST_AMT As Double = 0
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
                        GST = CGST_RATE + SGST_RATE
                        chkrs3.Close()
                        GST_AMT = GST * amt / 100

                        Dim net As Double
                        Dim rnd As Integer
                        rnd = GST_AMT - Math.Round(GST_AMT)
                        If rnd >= 50 Then
                            GST_AMT = Math.Round(GST_AMT) + 1
                        Else
                            GST_AMT = Math.Round(GST_AMT)
                        End If

                        net = amt + GST_AMT
                        CGST_TAXAMT = GST_AMT / 2


                        'CGST_TAXAMT = amt * CGST_RATE / 100
                        CGST_TAXAMT = Math.Round(GST_AMT / 2, 2, MidpointRounding.AwayFromZero)
                        'SGST_TAXAMT = amt * SGST_RATE / 100
                        SGST_TAXAMT = Math.Round(GST_AMT / 2, 2, MidpointRounding.AwayFromZero)

                        payable = amt + CGST_TAXAMT + SGST_TAXAMT
                        Label20.Text = amt.ToString & " + " & GST_AMT.ToString
                        Label19.Text = Format(payable, "0.00")      'Round(Label1.Caption, 2)
                        RentComboBox.DisplayMember = "Rent"
                        RentComboBox.ValueMember = "Month"
                        Dim tb As New DataTable
                        tb.Columns.Add("Month", GetType(Integer))
                        tb.Columns.Add("Rent", GetType(String))
                        For i As Integer = 1 To 24
                            If DataGridView2.RowCount >= 1 Then
                                tb.Rows.Add(i, Convert.ToDateTime(DataGridView2.Item(5, (DataGridView2.RowCount - 1)).Value).AddMonths(i) & " - " & (Convert.ToDouble(DataGridView2.Item(17, DataGridView2.RowCount - 1).Value) + (Convert.ToDouble(DataGridView2.Item(11, 0).Value) * i)).ToString)
                            Else
                                tb.Rows.Add(i, DateTime.Now.AddMonths(i).ToShortDateString & " - " & (payable * (i + 1)))
                            End If

                            '  
                            ' tb.Rows.Add(i, DateTime.Now.AddMonths(i).ToShortDateString & " - " & ((Convert.ToDouble(DataGridView2.Item(11, 0).Value) * i).ToString))
                        Next
                        RentComboBox.DataSource = tb
                        xcon.Close()
                    End If

                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error against bill: " & ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = getinvoicesr()
        GroupBox1.BringToFront()
        ' GroupBox1.Visible = True
        GroupBox2.SendToBack()
        Label17.Text = "Receipt Detail"
        indexorder = "GODWN_NO"
    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Dim errorMsg As String
        'If GrpAddCorrect = "A" And
        CheckBox1.Checked = False
        If bValidateamount = True And GrpAddCorrect <> "" Then
            Dim adjamt As Double = 0
            Dim adjable As Double = 0

            Dim destination_abroad As Boolean
            Dim checkedcount = 0
            '  MsgBox(Label19.Text
            If TextBox1.Text.Trim() = "" Then
                TextBox1.Text = "0.00"
            End If
            payable = Label19.Text
            For X As Integer = 0 To DataGridView2.RowCount - 1
                destination_abroad = Convert.ToBoolean(DataGridView2.Item(0, X).EditedFormattedValue)
                If destination_abroad = True Then
                    adjamt = adjamt + Convert.ToDouble(DataGridView2.Item(11, X).Value)
                    payable = Convert.ToDouble(DataGridView2.Item(11, X).Value)
                    checkedcount = checkedcount + 1
                End If
                adjable = adjable + Convert.ToDouble(DataGridView2.Item(11, X).Value)
                ' adjamt = adjamt + Convert.ToDouble(DataGridView2.Item(11, X).Value)
                ' payable = Convert.ToDouble(DataGridView2.Item(11, X).Value)
            Next
            If checkedcount < DataGridView2.RowCount And (Convert.ToDouble(TextBox1.Text) - adjamt) > payable Then
                errorMsg = "Please select bill for adjustment..."
                e.Cancel = True
                Me.ErrorProvider1.SetError(DataGridView2, errorMsg)
            Else
                If Convert.ToDouble(TextBox1.Text) <> adjamt Then
                    If Convert.ToDouble(TextBox1.Text) > adjable Then
                        Dim dba As Double = Convert.ToDouble(TextBox1.Text) Mod payable
                        Dim dbi As Integer = (Convert.ToDouble(TextBox1.Text) - adjable) Mod payable
                        If dba > 0 Then
                            errorMsg = "Part advance payment is not accepted..."
                            e.Cancel = True
                            TextBox1.Select(0, TextBox1.Text.Length)

                            ' Set the ErrorProvider error with the text to display. 
                            Me.ErrorProvider1.SetError(TextBox1, errorMsg)
                        Else
                            CheckBox1.Checked = True
                            Exit Sub
                        End If
                    Else
                        errorMsg = "Part Payment is not accepted..."
                        e.Cancel = True
                        TextBox1.Select(0, TextBox1.Text.Length)
                    End If
                    ' Set the ErrorProvider error with the text to display. 
                    Me.ErrorProvider1.SetError(TextBox1, errorMsg)
                Else
                    If TextBox1.Text.Equals("0.00") Or TextBox1.Text.Trim.Equals("") Then
                        errorMsg = "Enter Amount..."
                        e.Cancel = True
                        TextBox1.Select(0, TextBox1.Text.Length)

                        ' Set the ErrorProvider error with the text to display. 
                        Me.ErrorProvider1.SetError(TextBox1, errorMsg)
                    End If
                End If
            End If

        End If
        ' get_display_data()
    End Sub

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(DataGridView2, "")
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If GrpAddCorrect <> "" Then
            If RadioButton1.Checked = True Then
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox5.Enabled = True
                TextBox6.Enabled = True
                DateTimePicker2.Enabled = True
                Label12.Text = "Bank Account Detail"
            Else
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox6.Enabled = True
                Label12.Text = "Paid By (with Phone Number)"
                DateTimePicker2.Enabled = False
            End If
        End If
    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try
            ''''''''''''''''''check if newer receipt is exist'''''''''
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            Dim str As String = "SELECT * FROM RECEIPT where REC_NO >" & Convert.ToInt16(TextBox2.Text) & " AND REC_DATE>FORMAT('" & DateTimePicker1.Value & "','DD/MM/YY') And [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "'"
            chkrs1.Open("SELECT * FROM RECEIPT where REC_NO >" & Convert.ToInt16(TextBox2.Text) & " AND REC_DATE>FORMAT('" & DateTimePicker1.Value & "','DD/MM/YY') And [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "'", xcon)
            Do While chkrs1.EOF = False
                MsgBox("You can edit only latest receipt")
                Exit Sub
            Loop
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            GrpAddCorrect = "C"
            Label22.Text = "Edit"
            DataGridView2.Enabled = True
            DataGridView1.Enabled = False
            'Datprimaryrs.Recordset.Edit
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            TextBox2.Enabled = False
            TextBox1.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            ComboBox3.Enabled = False
            ComboBox4.Enabled = False
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RentComboBox.Enabled = True
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            ' CheckBox1.Checked = False
            Dim i As Integer = DataGridView1.CurrentRow.Index
            fillgrid2(DataGridView1.Item(1, i).Value, DataGridView1.Item(2, i).Value, DataGridView1.Item(4, i).Value, DataGridView1.Item(3, i).Value)
            rownum = DataGridView1.CurrentRow.Index

            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If ValidateChildren() Then
            insertData()
            DataGridView2.Enabled = False
            DataGridView1.Enabled = True
            rentsuggestion = False
            If DataGridView1.RowCount >= 1 Then
                rownum = DataGridView1.RowCount - 1
            End If
            If GrpAddCorrect = "C" Then
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                disablefields()
                DataGridView1.Rows(rownum).Selected = True
            Else
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                disablefields()
                rownum = DataGridView1.RowCount - 1
                DataGridView1.Rows(rownum).Selected = True
            End If
            Label22.Text = "View"
            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub insertData()
        Dim save As String
        Dim savedetails As String
        Dim stus As String
        Dim transaction As OleDbTransaction
        If RadioButton1.Checked Then
            stus = "Q"
        Else
            stus = "C"
        End If
        Dim DETAMT As Double = 0
        Dim ADJAMTT As Double = 0
        Dim tmpdate As String = Convert.ToDateTime(DateTimePicker1.Value.ToString)
        Dim tmprcno As Integer = Convert.ToInt32(TextBox2.Text)
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
            If GrpAddCorrect = "C" Then
                '''''REVERSE UPDATE
                Dim a As String = "delete * from [RECEIPT] where REC_NO=" & Convert.ToInt32(TextBox2.Text) & " AND year(REC_DATE)='" & Convert.ToDateTime(DateTimePicker1.Value.ToString).Year & "'"
                For i = 0 To DataGridView2.RowCount - 1

                    save = "UPDATE [BILL] SET REC_NO=Null, REC_DATE=Null WHERE INVOICE_NO='" & DataGridView2.Item(1, i).Value & "' AND BILL_DATE=#" & Convert.ToDateTime(DataGridView2.Item(5, i).Value) & "#" ' sorry about that
                    objcmd.CommandText = save
                    objcmd.ExecuteNonQuery()

                    objcmd.CommandText = "delete * from [RECIPTBILL] where REC_NO=" & Convert.ToInt32(TextBox2.Text) & " AND INVOICE_NO ='" & DataGridView2.Item(1, i).Value & "'"
                    objcmd.ExecuteNonQuery()
                    ' objcmd.Dispose()
                Next

                objcmd.CommandText = "delete * from [RECEIPT] where REC_NO=" & Convert.ToInt32(TextBox2.Text) & " AND year(REC_DATE)='" & Convert.ToDateTime(DateTimePicker1.Value.ToString).Year & "'"
                objcmd.ExecuteNonQuery()

                For i = 0 To DataGridView2.RowCount - 1
                    If Not IsDBNull(DataGridView2.Item(0, i).Value) Then


                        If DataGridView2.Item(0, i).Value = 1 Then
                            objcmd.CommandText = "UPDATE [BILL] SET REC_NO='" & Convert.ToInt32(TextBox2.Text) & "', REC_DATE ='" & DateTimePicker1.Value.ToShortDateString & "' WHERE INVOICE_NO='" & DataGridView2.Item(1, i).Value & "' AND BILL_DATE=#" & DataGridView2.Item(5, i).Value & "#" ' sorry about that
                            objcmd.ExecuteNonQuery()
                            objcmd.CommandText = "INSERT INTO [RECIPTBILL](REC_NO,INVOICE_NO,AMT,REC_DATE) VALUES('" & Convert.ToInt32(TextBox2.Text) & "','" & DataGridView2.Item(1, i).Value & "','" & Convert.ToDouble(DataGridView2.Item(11, i).Value) & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "')"
                            objcmd.ExecuteNonQuery()
                            DETAMT = DETAMT + Convert.ToDouble(DataGridView2.Item(11, i).Value)
                        End If
                    End If
                Next i
                ADJAMTT = Convert.ToDouble(TextBox1.Text) - DETAMT
                objcmd.CommandText = "INSERT INTO [RECEIPT]([GROUP],GODWN_NO,REC_DATE,REC_NO,AMOUNT,ADVANCE,CASH_CHEQUE,BANK_NAME,BRANCH,CHEQUE_NUMBER,CHEQUE_DATE,BANK_AC_DETAIL,ADJ_AMT) VALUES('" & ComboBox3.SelectedValue.ToString & "','" & ComboBox4.SelectedValue.ToString & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & Convert.ToInt32(TextBox2.Text) & "','" & Convert.ToDouble(TextBox1.Text) & "'," & CheckBox1.Checked & ",'" & stus & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox3.Text & "','" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "','" & TextBox6.Text & "','" & ADJAMTT & "')"
                objcmd.ExecuteNonQuery()

            Else

                For i = 0 To DataGridView2.RowCount - 1
                    If Not IsDBNull(DataGridView2.Item(0, i).Value) Then
                        If DataGridView2.Item(0, i).Value = 1 Then
                            objcmd.CommandText = "UPDATE [BILL] SET REC_NO='" & Convert.ToInt32(TextBox2.Text) & "', REC_DATE ='" & DateTimePicker1.Value.ToShortDateString & "' WHERE INVOICE_NO='" & DataGridView2.Item(1, i).Value & "' AND BILL_DATE=#" & DataGridView2.Item(5, i).Value & "#" ' sorry about that
                            objcmd.ExecuteNonQuery()
                            objcmd.CommandText = "INSERT INTO [RECIPTBILL](REC_NO,INVOICE_NO,AMT,REC_DATE) VALUES('" & Convert.ToInt32(TextBox2.Text) & "','" & DataGridView2.Item(1, i).Value & "','" & Convert.ToDouble(DataGridView2.Item(11, i).Value) & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "')"
                            objcmd.ExecuteNonQuery()
                            DETAMT = DETAMT + Convert.ToDouble(DataGridView2.Item(11, i).Value)
                        End If
                    End If
                Next i
                ADJAMTT = Convert.ToDouble(TextBox1.Text) - DETAMT
                objcmd.CommandText = "INSERT INTO [RECEIPT]([GROUP],GODWN_NO,REC_DATE,REC_NO,AMOUNT,ADVANCE,CASH_CHEQUE,BANK_NAME,BRANCH,CHEQUE_NUMBER,CHEQUE_DATE,BANK_AC_DETAIL,ADJ_AMT) VALUES('" & ComboBox3.SelectedValue.ToString & "','" & ComboBox4.SelectedValue.ToString & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & Convert.ToInt32(TextBox2.Text) & "','" & Convert.ToDouble(TextBox1.Text) & "'," & CheckBox1.Checked & ",'" & stus & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox3.Text & "','" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "','" & TextBox6.Text & "','" & ADJAMTT & "')"
                objcmd.ExecuteNonQuery()
            End If
            DataGridView1.Update()
            transaction.Commit()
            MsgBox("Data Inserted successfully in database", vbInformation)
        Catch ex As Exception
            MsgBox("Exception: Data Insertion in RECEIPT table in database" & ex.Message)
            Try
                transaction.Rollback()

            Catch
                ' Do nothing here; transaction is not active.
            End Try
        End Try
        '  GrpAddCorrect = ""  '''dipti lat change
        ShowData()

    End Sub
    Private Sub doSQL(ByVal sql As String, ByVal transaction As OleDbTransaction)
        Dim objcmd As New OleDb.OleDbCommand
        objcmd.Connection = MyConn
        objcmd.Transaction = transaction
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sql
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If DataGridView2.CurrentCell.ColumnIndex = 0 Then
            Dim current_abroad = Convert.ToBoolean(DataGridView2.CurrentCell.EditedFormattedValue)
            If current_abroad = True Then
                For X As Integer = 0 To DataGridView2.CurrentRow.Index
                    DataGridView2.Item(0, X).Value = True
                Next
            Else
                For X As Integer = DataGridView2.CurrentRow.Index To DataGridView2.RowCount - 1
                    DataGridView2.Item(0, X).Value = False
                Next
            End If



            Dim destination_abroad As Boolean
            Dim TMPAMT As Double
            TextBox1.Text = 0
            For X As Integer = 0 To DataGridView2.RowCount - 1
                destination_abroad = Convert.ToBoolean(DataGridView2.Item(0, X).EditedFormattedValue)
                If destination_abroad = True Then
                    TMPAMT = Convert.ToDouble(TextBox1.Text)
                    TextBox1.Text = Format(CSng(TMPAMT + Convert.ToDouble(DataGridView2.Item(11, X).Value)), "#####0.00")     ' TMPAMT + Convert.ToDouble(DataGridView2.Item(10, DataGridView2.CurrentRow.Index).Value)
                End If
            Next

        End If
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        ErrorProvider1.SetError(TextBox2, "")
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If GrpAddCorrect = "A" And bValidateinvoice = True And GrpAddCorrect <> "" Then
            Dim adjamt As Double = 0
            Dim payable As Double = 0
            If TextBox2.Text.Equals("") Then
                Dim errorMsg As String = "Please insert receipt number..."
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox2, errorMsg)
            Else
                Try

                    Dim RECFOUND As Boolean
                    If xcon.State = ConnectionState.Open Then
                    Else
                        xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                    End If
                    Dim iDate As String
                    Dim fDate As DateTime
                    Dim oDate As String
                    Dim foDate As DateTime
                    If (Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) >= 4) Then
                        iDate = "30/04/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                        fDate = Convert.ToDateTime(iDate)
                        oDate = "31/03/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) + 1))
                        foDate = Convert.ToDateTime(oDate)
                    Else
                        iDate = "30/04/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) - 1))
                        fDate = Convert.ToDateTime(iDate)
                        oDate = "31/03/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                        foDate = Convert.ToDateTime(oDate)
                    End If


                    '  Dim SRD As String = "SELECT [RECEIPT].* from [RECEIPT] where [GROUP]='" & ComboBox3.Text & "' AND GODWN_NO='" & ComboBox4.Text & "' AND Year([RECEIPT].REC_DATE)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [RECEIPT].REC_NO"
                    Dim strn As String = "Select * FROM BILL where [BILL].bill_date>=Format('" & fDate & "', 'dd/mm/yyyy') and [BILL].bill_date<=Format('" & foDate & "', 'dd/mm/yyyy') order by Year([BILL].bill_date)+[BILL].INVOICE_NO"
                    chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [receipt].rec_date>=Format('" & fDate & "', 'dd/mm/yyyy') and [receipt].rec_date<=Format('" & foDate & "', 'dd/mm/yyyy') AND REC_NO=" & Convert.ToInt32(TextBox2.Text) & " order by [RECEIPT].REC_NO", xcon)
                    ' chkrs1.MoveLast()
                    Do While chkrs1.EOF = False
                        RECFOUND = True
                        chkrs1.MoveNext()
                    Loop
                    chkrs1.Close()
                    xcon.Close()
                    If RECFOUND Then
                        Dim errorMsg As String = "Duplicate receipt number not allowed..."
                        e.Cancel = True
                        TextBox2.Select(0, TextBox2.Text.Length)
                        ' Set the ErrorProvider error with the text to display. 
                        Me.ErrorProvider1.SetError(TextBox2, errorMsg)
                    End If
                Catch ex As Exception
                    MsgBox("Exception: duplicate receipt check :" & ex.Message)
                End Try

            End If

        End If
    End Sub
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidategodown = False
        bValidatetype = False
        bValidateinvoice = False
        bValidatedate = False
        bValidateamount = False
    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidategodown = True
        bValidatetype = True
        bValidateinvoice = True
        bValidatedate = True
        bValidateamount = True
    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            GrpAddCorrect = ""
            rentsuggestion = False
            Label22.Text = "View"
            ErrorProvider1.Clear()
            DataGridView2.Enabled = False
            DataGridView1.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            disablefields()
            navigateenable()
            Label14.Text = ""
            Label15.Text = ""
            If DataGridView2.RowCount > 1 Then
                DataGridView2.DataSource = ""
            End If
            ShowData()
            LodaDataToTextBox()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    'Private Sub ComboBox3_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox3.Validating

    '    If Application.OpenForms().OfType(Of GodownHelp).Any And ComboBox3.Text.Trim.Equals("") Then
    '        e.Cancel = True
    '        bValidatetype = False
    '        ComboBox4.Select(0, ComboBox4.Text.Length)
    '    Else
    '        Dim errorMsg As String = "Please select godown type"
    '        If GrpAddCorrect = "A" And bValidatetype = True And ComboBox3.Text.Trim.Equals("") Then
    '            ' Cancel the event and select the text to be corrected by the user.
    '            e.Cancel = True
    '            ComboBox3.Select(0, ComboBox3.Text.Length)
    '            ' Set the ErrorProvider error with the text to display. 
    '            Me.ErrorProvider1.SetError(ComboBox3, errorMsg)
    '        End If
    '    End If
    'End Sub
    'Private Sub ComboBox3_Validated(sender As Object, e As EventArgs) Handles ComboBox3.Validated
    '    If Application.OpenForms().OfType(Of GodownHelp).Any Then

    '    Else
    '        ErrorProvider1.SetError(ComboBox3, "")
    '    End If
    'End Sub
    Private Sub ComboBox4_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox4.Validating
        If Application.OpenForms().OfType(Of GodownHelp).Any Then
        Else

            If GrpAddCorrect = "A" And bValidatetype = True And ComboBox3.Text.Trim.Equals("") Then
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                ComboBox3.Select(0, ComboBox3.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                MsgBox("Please select Godown Type")
                Exit Sub
            End If
            Dim errorMsg As String = "Please select godown number"
            If GrpAddCorrect = "A" And bValidategodown = True And ComboBox4.Text.Trim.Equals("") Then
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                ComboBox4.Select(0, ComboBox4.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(ComboBox4, errorMsg)
            End If
        End If
    End Sub

    Private Sub ComboBox4_Validated(sender As Object, e As EventArgs) Handles ComboBox4.Validated
        ErrorProvider1.SetError(ComboBox4, "")
    End Sub

    Private Sub TextBox3_Validated(sender As Object, e As EventArgs) Handles TextBox3.Validated
        ErrorProvider1.SetError(TextBox3, "")
    End Sub

    Private Sub TextBox3_Validating(sender As Object, e As CancelEventArgs) Handles TextBox3.Validating
        If GrpAddCorrect = "A" And bValidateinvoice = True And GrpAddCorrect <> "" Then
            If RadioButton1.Checked And TextBox3.Text.Equals("") Then
                Dim errorMsg As String = "Please insert cheque number..."
                e.Cancel = True
                TextBox3.Select(0, TextBox3.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox3, errorMsg)
            End If

        End If
    End Sub
    Private Sub TextBox4_Validated(sender As Object, e As EventArgs) Handles TextBox4.Validated
        ErrorProvider1.SetError(TextBox4, "")
    End Sub

    Private Sub TextBox4_Validating(sender As Object, e As CancelEventArgs) Handles TextBox4.Validating
        If GrpAddCorrect = "A" And bValidateinvoice = True And GrpAddCorrect <> "" Then
            If RadioButton1.Checked And TextBox4.Text.Equals("") Then
                Dim errorMsg As String = "Please insert bank name..."
                e.Cancel = True
                TextBox4.Select(0, TextBox4.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox4, errorMsg)
            End If

        End If
    End Sub
    Private Sub TextBox5_Validated(sender As Object, e As EventArgs) Handles TextBox5.Validated
        ErrorProvider1.SetError(TextBox5, "")
    End Sub

    Private Sub TextBox5_Validating(sender As Object, e As CancelEventArgs) Handles TextBox5.Validating
        If GrpAddCorrect = "A" And bValidateinvoice = True And GrpAddCorrect <> "" Then
            If RadioButton1.Checked And TextBox5.Text.Equals("") Then
                Dim errorMsg As String = "Please insert branch..."
                e.Cancel = True
                TextBox5.Select(0, TextBox5.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox5, errorMsg)
            End If

        End If
    End Sub
    Private Sub TextBox6_Validated(sender As Object, e As EventArgs) Handles TextBox6.Validated
        ErrorProvider1.SetError(TextBox6, "")
    End Sub

    Private Sub TextBox6_Validating(sender As Object, e As CancelEventArgs) Handles TextBox6.Validating
        If GrpAddCorrect = "A" And bValidateinvoice = True And GrpAddCorrect <> "" Then
            If RadioButton1.Checked And TextBox6.Text.Equals("") Then
                Dim errorMsg As String = "Please insert bank account detail..."
                e.Cancel = True
                TextBox6.Select(0, TextBox6.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox6, errorMsg)
            End If
            If RadioButton2.Checked And TextBox6.Text.Equals("") Then
                Dim errorMsg As String = "Please insert contact person & phone number..."
                e.Cancel = True
                TextBox6.Select(0, TextBox6.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox6, errorMsg)
            End If
        End If
    End Sub

    Private Sub DataGridView3_LostFocus(sender As Object, e As EventArgs) Handles DataGridView3.LostFocus
        'GroupBox1.BringToFront()
        '' GroupBox1.Visible = True
        'GroupBox2.SendToBack()
        'Label17.Text = "Receipt Detail"
    End Sub

    Private Sub GroupBox2_LostFocus(sender As Object, e As EventArgs) Handles GroupBox2.LostFocus
        'GroupBox1.BringToFront()
        '' GroupBox1.Visible = True
        'GroupBox2.SendToBack()
        'Label17.Text = "Receipt Detail"
        'indexorder = "GODWN_NO"
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.Select(0, TextBox1.Text.Length)
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        ' da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where " & indexorder & " Like '%" & TxtSrch.Text & "%' order by [RECEIPT].REC_NO", MyConn)
        da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].*,[S].[P_CODE] from [RECEIPT] LEFT OUTER JOIN (SELECT [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE] FROM [BILL] GROUP BY [BILL].[REC_DATE],[BILL].[REC_NO],[BILL].[P_CODE]) AS S ON [RECEIPT].[REC_DATE]=[S].[REC_DATE] AND TRIM(STR([RECEIPT].[REC_NO]))=[S].[REC_NO] where " & indexorder & " Like '%" & TxtSrch.Text & "%'  order by [RECEIPT].[REC_DATE],[RECEIPT].REC_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

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
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        DataGridView1.Enabled = False
        MyConn = New OleDbConnection(connString)
        Dim transaction As OleDbTransaction
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        Dim tmpdldate As Date = Convert.ToDateTime(DateTimePicker1.Value.ToString).AddMonths(1)
        Dim TMPDLgR As String = ComboBox3.Text
        Dim tmpgd As String = ComboBox4.Text
        da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where REC_DATE=#" & tmpdldate & "# AND [GROUP]='" & TMPDLgR & "' AND GODWN_NO='" & tmpgd & "' order by [RECEIPT].REC_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "RECEIPT")

        If ds.Tables(0).Rows.Count > 0 Then
            MsgBox("You can delete latest receipt only... Delete latest receipt first..")
            DataGridView2.Enabled = True
            Exit Sub
        End If
        da.Dispose()
        ds.Dispose()
        ' MyConn.Close() ' close connection
        Try
            Dim kk As Integer = MsgBox("[" & Trim(TextBox2.Text) & "]  Delete Record ?", vbYesNo + vbDefaultButton2)
            If kk = 6 Then  'i assume yes

                Dim objcmd As New OleDb.OleDbCommand
                Try
                    transaction = MyConn.BeginTransaction(IsolationLevel.ReadCommitted)
                    objcmd.Connection = MyConn
                    objcmd.Transaction = transaction
                    objcmd.CommandType = CommandType.Text
                    objcmd.CommandText = "delete from [RECEIPT] where year(REC_DATE)='" & Convert.ToDateTime(DateTimePicker1.Value.ToString).Year & "' AND REC_NO=" & Convert.ToInt32(TextBox2.Text)
                    objcmd.ExecuteNonQuery()
                    '  MsgBox("Data deleted successfully from GODOWN table in database", vbInformation)

                    objcmd.CommandText = "DELETE FROM [RECIPTBILL] WHERE year(REC_DATE)='" & Convert.ToDateTime(DateTimePicker1.Value.ToString).Year & "' AND REC_NO=" & Convert.ToInt32(TextBox2.Text)
                    objcmd.ExecuteNonQuery()
                    '  MsgBox("Data deleted successfully from GODOWN table In database", vbInformation)
                    objcmd.CommandText = "UPDATE [BILL] Set REC_NO=Null, REC_DATE =Null WHERE REC_NO='" & Convert.ToInt32(TextBox2.Text) & "' AND year(REC_DATE)='" & Convert.ToDateTime(DateTimePicker1.Value.ToString).Year & "'"
                    objcmd.ExecuteNonQuery()
                    ShowData()
                    rownum = 0
                    LodaDataToTextBox()
                    transaction.Commit()

                    MsgBox("Receipt deleted successfully from database", vbInformation)
                    DataGridView1.Enabled = True
                Catch ex As Exception
                    MsgBox("Exception: Data Delete module " & ex.Message)
                    Try
                        transaction.Rollback()

                    Catch
                        ' Do nothing here; transaction is not active.
                    End Try
                End Try
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Delete module: " & ex.Message)
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If GrpAddCorrect = "A" Then
            lastdate = DateTimePicker1.Value
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True

        End If
    End Sub

    Private Sub RentComboBox_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles RentComboBox.SelectionChangeCommitted
        'Dim vl As Double
        ''Console.WriteLine(RentComboBox.SelectedText)
        'If Double.TryParse(RentComboBox.SelectedText.ToString.Substring(12), vl) Then
        '    TextBox1.Text = CDbl(Val(RentComboBox.SelectedText.ToString.Substring(12)))
        'End If
        rentsuggestion = True
    End Sub

    Private Sub RentComboBox_TextUpdate(sender As Object, e As EventArgs) Handles RentComboBox.TextUpdate
        'If RentComboBox.FindString(RentComboBox.Text) < 0 Then
        '    RentComboBox.Text = RentComboBox.Text.Remove(RentComboBox.Text.Length - 1)
        '    RentComboBox.SelectionStart = RentComboBox.Text.Length

        'End If
    End Sub

    Private Sub RentComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RentComboBox.SelectedIndexChanged

        If rentsuggestion = True Then
            DataGridView2.Columns(0).ReadOnly = False
            For X As Integer = 0 To DataGridView2.RowCount - 1
                ' Console.WriteLine("grid " & DataGridView2.Item(5, X).Value)
                If (Format(DataGridView2.Item(5, X).Value, "dd/MM/yyyy") <= Format(RentComboBox.Text.ToString().Substring(0, 10), "dd/MM/yyyy")) Then
                    Console.WriteLine(Format(DataGridView2.Item(5, X).Value, "dd/MM/yyyy"))

                    If DataGridView2.Item(0, X).Value = False Then

                        DataGridView2.CurrentCell = DataGridView2.Item(0, X)
                        DataGridView2.BeginEdit(True)
                        DataGridView2.Item(0, X).Value = True
                        DataGridView2.RefreshEdit()

                    End If
                End If
                ' Console.WriteLine("combo " & RentComboBox.Text.ToString().Substring(0, 10))

            Next
            TextBox1.Text = RentComboBox.Text.ToString().Substring(12)
            TextBox1.Focus()


        End If
    End Sub


End Class