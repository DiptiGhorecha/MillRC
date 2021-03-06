﻿Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
''' <summary>
''' Tables used - bill,godown,part,rent
''' when bulk invoices generated and after that user need invoice for any user
''' this module is used. When user select invoice date, first it checks that bulk invoices generated or not. Allow to proceed only if bulk invoice generated.
''' When user select group and godown number, get invoice number in serial itself and fill all other form details. When user click update button invoice is generated
''' at /invoices/pdf/<year>/<month> directory
''' </summary>

Public Class FrmInvSingle
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
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim dagp As OleDbDataAdapter
    Dim dsgp As DataSet
    Dim indexorder As String = "GODWN_NO"   ''''variable to store sorting order field for datagrid
    Dim GrpAddCorrect As String
    Private bValidateinvoice As Boolean = True
    Private bValidatetype As Boolean = True
    Private bValidategodown As Boolean = True
    Private bValidatedate As Boolean = True
    Dim formloaded As Boolean = False
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim save As String
    Dim pdfpath As String
    Dim strReportFilePath As String
    Public FILE_NO As String
    Public cmdClicked As Boolean = False
    Public gridline As Integer = 0

    Private Sub FrmInvSingle_Load(sender As Object, e As EventArgs) Handles Me.Load
        '''''''set position of form 
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        '''''''enabled/disabled form components
        cmdAdd.Enabled = True
        cmdClose.Enabled = True
        cmdDelete.Enabled = True
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = False
        cmdCancel.Enabled = False
        disablefields()
        fillgroupcombo()   ''''''fill godown group code drop down
        fillgodowncombo()   '''''fill godown code drop down
        ShowData()          '''''load data into data grid view using bill table
        LodaDataToTextBox()   ''''' load data for current datagridview row to form
        formloaded = True
    End Sub
    Public Function fillgodowncombo()
        '''''fill godown combo box with godown number for current godowns using godown table
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
        ''''''fill godown group combo box using group table
        Try
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

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        ErrorProvider1.SetError(TextBox2, "")
    End Sub
    Function disablefields()
        ''''''''disable form elements
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ChkLogo.Enabled = False
        DateTimePicker1.Enabled = False
    End Function
    Private Sub ShowData()
        ''''show invoices data into data grid view using bill table
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
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
            DataGridView2.Columns(15).Visible = True
            DataGridView2.Columns(1).Visible = True
            DataGridView2.Columns(0).HeaderText = "Invoice No."
            DataGridView2.Columns(1).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 120
            DataGridView2.Columns(1).Width = 51
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(15).Width = 300
            DataGridView2.Columns(2).HeaderText = "Godown"
            DataGridView2.Columns(15).HeaderText = "Tenant"
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
            DataGridView2.Columns(13).Visible = False
            DataGridView2.Columns(14).Visible = False
            DataGridView2.Columns(4).Width = 80

            Label21.Text = "Total : " & DataGridView2.RowCount - 1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub LodaDataToTextBox()
        '''''''''load data from currently selected datagrid view - bill table used to bind data with grid
        Try
            Dim i As Integer
            TextBox1.Text = ""         ''''' godown number
            TextBox2.Text = ""         ''''invoice number
            TextBox3.Text = ""         '''''' gst number
            TextBox4.Text = ""         '''''' bill amount
            TextBox5.Text = ""         '''''taxable amount
            TextBox6.Text = ""         ''''cgst amount
            TextBox7.Text = ""         ''''sgst_amt
            TextBox8.Text = ""         ''''net amount
            TextBox9.Text = ""         ''''cgst rate
            TextBox10.Text = ""        ''''sgst rate
            TextBox11.Text = ""        ''''email id for tenant
            TextBox12.Text = ""        '''''gst code for tenant
            TextBox13.Text = ""        ''''reference number
            ComboBox1.Text = ""      ''''from month
            ComboBox2.Text = ""      '''' to month
            ComboBox3.Text = ""      ''''godown group
            ComboBox4.Text = ""      '''' godown number
            DateTimePicker1.Text = ""     ''''invoice date
            If DataGridView2.RowCount > 1 Then
                If gridline > 0 Then
                    i = gridline
                Else
                    i = DataGridView2.CurrentRow.Index

                End If

                gridline = i
                If Not IsDBNull(DataGridView2.Item(4, i).Value) Then
                    DateTimePicker1.Value = GetValue(DataGridView2.Item(4, i).Value)    '''''''bill_date from bill table
                    ComboBox1.Text = DateTimePicker1.Value.Month
                    ComboBox2.Text = DateTimePicker1.Value.Year
                End If
                If Not IsDBNull(DataGridView2.Item(1, i).Value) Then
                    ComboBox3.Text = GetValue(DataGridView2.Item(1, i).Value)           '''''group from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(2, i).Value) Then
                    TextBox1.Text = GetValue(DataGridView2.Item(2, i).Value)            '''''''godwn_no from bill
                    ComboBox4.Text = GetValue(DataGridView2.Item(2, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(0, i).Value) Then
                    TextBox2.Text = GetValue(DataGridView2.Item(0, i).Value)             '''''invoice_no from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(11, i).Value) Then
                    TextBox3.Text = GetValue(DataGridView2.Item(11, i).Value)              '''HSN from bill table
                    If TextBox3.Text = "997212" Then
                        Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"   '''''gst description
                    Else
                        Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                    End If
                End If
                If Not IsDBNull(DataGridView2.Item(15, i).Value) Then
                    Label13.Text = GetValue(DataGridView2.Item(15, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                    TextBox4.Text = Format(CSng(GetValue(DataGridView2.Item(5, i).Value)), "#####0.00")        ''''''bill_amt from bill table
                    TextBox5.Text = Format(CSng(GetValue(DataGridView2.Item(5, i).Value)), "#####0.00")      ''''''bill_amt from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(7, i).Value) Then
                    TextBox6.Text = Format(CSng(GetValue(DataGridView2.Item(7, i).Value)), "#####0.00")        '''''''cgst_amt from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(9, i).Value) Then
                    TextBox7.Text = Format(CSng(GetValue(DataGridView2.Item(9, i).Value)), "#####0.00")        ''''''''sgst_amt from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(10, i).Value) Then
                    TextBox8.Text = Format(CSng(GetValue(DataGridView2.Item(10, i).Value)), "#####0.00")       '''net_amt from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(6, i).Value) Then
                    TextBox9.Text = GetValue(DataGridView2.Item(6, i).Value)              '''''cgst_rate from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(8, i).Value) Then
                    TextBox10.Text = GetValue(DataGridView2.Item(8, i).Value)             '''''sgst_rate from bill table
                End If
                If Not IsDBNull(DataGridView2.Item(12, i).Value) Then                    ''''''srno from bill table
                    TextBox13.Text = DataGridView2.Item(12, i).Value
                End If
                If xcon.State = ConnectionState.Open Then
                Else
                    xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                End If
                ''''''select tentant detail form party table for p_code of bill table and append it in label13
                chkrs4.Open("SELECT * FROM PARTY WHERE P_CODE='" & GetValue(DataGridView2.Item(3, i).Value) & "'", xcon)
                Dim TAD1 As String
                Dim TAD2 As String
                Dim TAD3 As String
                Dim TCITY As String
                Dim TSTATE As String
                If chkrs4.EOF = False Then

                    If IsDBNull(chkrs4.Fields(2).Value) Then
                        TAD1 = ""
                    Else
                        If Trim(chkrs4.Fields(2).Value).Equals("") Then
                            TAD1 = ""
                        Else
                            TAD1 = chkrs4.Fields(2).Value.replace("& vbLf & vbLf", "")
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(2).Value
                        End If
                    End If
                    If IsDBNull(chkrs4.Fields(3).Value) Then
                        TAD2 = ""
                    Else
                        If Trim(chkrs4.Fields(3).Value).Equals("") Then
                            TAD2 = ""
                        Else
                            TAD2 = chkrs4.Fields(3).Value
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(3).Value
                        End If
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                        TAD3 = ""
                    Else
                        If Trim(chkrs4.Fields(4).Value).Equals("") Then
                            TAD3 = ""
                        Else
                            TAD3 = chkrs4.Fields(4).Value
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(4).Value
                        End If
                    End If
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                        TCITY = ""
                    Else
                        If Trim(chkrs4.Fields(5).Value).Equals("") Then
                            TCITY = ""
                        Else
                            TCITY = chkrs4.Fields(5).Value
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(5).Value
                        End If
                    End If
                    If IsDBNull(chkrs4.Fields(17).Value) Then
                        TSTATE = ""
                    Else
                        If Trim(chkrs4.Fields(17).Value).Equals("") Then
                            TSTATE = ""
                        Else
                            TSTATE = chkrs4.Fields(17).Value
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(17).Value      ''''state from party table
                        End If
                    End If



                    If Not IsDBNull(chkrs4.Fields(19).Value) Then
                        TextBox12.Text = chkrs4.Fields(19).Value      '''''gst from party table
                    End If
                    If Not IsDBNull(chkrs4.Fields(18).Value) Then
                        TextBox11.Text = chkrs4.Fields(18).Value      ''''email_id from party table
                    End If
                End If
                chkrs4.Close()
                xcon.Close()
                Label21.Text = "Total : " & DataGridView2.RowCount - 1
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try




    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub KeyUpHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.F1 Then
        End If
    End Sub

    Private Sub KeyDownHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True

    End Sub

    Private Sub FrmInvSingle_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ''''''show help form for godowns on pressing F1 key
        If e.KeyCode = Keys.F1 Then
            DataGridView2.Visible = True
            GroupBox5.Visible = True
            GroupBox4.Visible = True
            Me.Width = Me.Width + DataGridView2.Width - 100
            Me.Height = Me.Height + 80
            ShowData()
        End If
    End Sub
    Private Sub cmdFirst_Click(sender As Object, e As EventArgs) Handles cmdFirst.Click
        '''''set 1st row as current row of data grid
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(0).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(0)
        '  rownum = 0
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        '''''set previous row as current row of data grid
        Dim intRow As Integer = DataGridView2.CurrentRow.Index
        If intRow > 0 Then
            DataGridView2.CurrentRow.Selected = False
            DataGridView2.Rows(intRow - 1).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(intRow - 1).Cells(0)
            '   rownum = intRow - 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        '''''set next row as current row of data grid
        Dim intRow As Integer = DataGridView2.CurrentRow.Index
        If intRow < DataGridView2.RowCount - 1 Then
            DataGridView2.CurrentRow.Selected = False
            DataGridView2.Rows(intRow + 1).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(intRow + 1).Cells(0)
            '   rownum = intRow + 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        '''''set last row as current row of data grid
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(DataGridView2.RowCount - 1).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(0)
        '   rownum = DataGridView2.RowCount - 1
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        '''''this event will fire when user write anything in search text box
        '''''It will filter data from gdtrans table and bind with datagrid view
        MyConn = New OleDbConnection(connString)
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [BILL].GROUP,[BILL].GODWN_NO,[BILL].BILL_DATE", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        ''''''''set index order for searching and search textbox label according to datagrid column user clicked 
        If e.ColumnIndex = 15 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        Else
            indexorder = "[GODOWN].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
        End If
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

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        '''''''''event associated with cancel button
        ''''made entry area, update button and cancel button disabled 
        ''''made grid view, navigation buttons , add button enabled so user can again select data from grid or using navigation buttons
        Try
            GrpAddCorrect = ""
            ErrorProvider1.Clear()
            DataGridView2.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            disablefields()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub
    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        '''''''load data to form for current row when user click on data grid
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyDown
        ''''''''when user select datagrid view row using up, down arrow key and press enter - show that row data in form
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView2_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyUp
        ''''''''when user select datagrid view row using up, down arrow key and press enter - show that row data in form
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        '''''''load data to form for current row when user double click on data grid
        LodaDataToTextBox()
    End Sub
    Public Function getinvoiceno()
        ''''to get invoice number in series 
        Try
            Dim INVNO As String
            Dim INVNOTMP As String
            Dim nom As Integer

            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            chkrs1.Open("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [BILL].INVOICE_NO", xcon)
            ' chkrs1.MoveLast()
            Do While chkrs1.EOF = False
                INVNO = chkrs1.Fields(0).Value
                chkrs1.MoveNext()
            Loop
            chkrs1.Close()
            xcon.Close()
            nom = Convert.ToInt32(INVNO) + 1
            INVNO = String.Format("{0:0000}", nom)
            Return INVNO
        Catch ex As Exception
            MsgBox("Exception: Get invoice No :" & ex.Message)
        End Try
    End Function
    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        '''''''''event associated with Add button
        ''''made entry area, update button and cancel button enabled
        ''''made grid view, navigation buttons , add button, search text box disabled so user can not again select data from grid or using navigation buttons

        Try
            GrpAddCorrect = "A"
            Label23.Text = "ADD"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            'textenable()
            TextBox2.Enabled = True
            TextBox2.Text = ""  'getinvoiceno()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            DateTimePicker1.Enabled = True
            ComboBox3.Enabled = True
            ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
            ComboBox1.Text = ""
            ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
            ComboBox2.Text = ""
            ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox3.Text = ""
            ComboBox4.Enabled = True
            ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
            ComboBox4.Text = ""
            ChkLogo.Enabled = True
            ChkLogo.Checked = False
            'ComboBox3.Select()
            DateTimePicker1.Value = Date.Today
            DateTimePicker1.Focus()
            Label13.Text = ""
            Label18.Text = ""
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub
    Private Sub navigatedisable()
        ''''disable navigation
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
        TxtSrch.Enabled = False
    End Sub

    Private Sub navigateenable()
        ''''enable navigation buttons
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        TxtSrch.Enabled = True
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        '''''''''auto generate invoice number get from function when invoice number text field get focus
        If DataGridView2.RowCount > 1 And GrpAddCorrect = "A" Then
            TextBox2.Text = getinvoiceno()
        End If
    End Sub
    Private Sub ComboBox4_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox4.Validating
        '''''check if invoice is already generated for selected month, year and godown number

        Dim errorMsg As String = "Please Select godown Number"
        If bValidategodown = True Then
            If ComboBox4.Text.Trim.Equals("") Then
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                ComboBox4.Select(0, ComboBox4.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(ComboBox4, errorMsg)
            Else
                MyConn = New OleDbConnection(connString)
                If MyConn.State = ConnectionState.Closed Then
                    MyConn.Open()
                End If

                Dim str As String = "SELECT [BILL].INVOICE_NO, [BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE from [BILL] where [GROUP]='" & Trim(ComboBox3.SelectedValue.ToString) & "' AND [GODWN_NO]='" & Trim(ComboBox4.SelectedValue.ToString) & "' AND BILL_DATE=FORMAT('" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','DD/MM/YYYY')"

                dag = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE from [BILL] where [GROUP]='" & Trim(ComboBox3.SelectedValue.ToString) & "' AND [GODWN_NO]='" & Trim(ComboBox4.SelectedValue.ToString) & "' AND BILL_DATE=FORMAT('" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','DD/MM/YYYY')", MyConn)
                dsg = New DataSet
                dsg.Clear()
                dag.Fill(dsg, "BILL")

                If dsg.Tables(0).Rows.Count > 0 And GrpAddCorrect <> "C" Then
                    errorMsg = "Invoice is already generated..."
                    e.Cancel = True
                    ComboBox4.Select(0, ComboBox4.Text.Length)
                    ' Set the ErrorProvider error with the text to display. 
                    Me.ErrorProvider1.SetError(ComboBox4, errorMsg)
                End If
                dag.Dispose()
                dsg.Dispose()
                MyConn.Close() ' close connection
            End If
        End If

    End Sub

    Private Sub ComboBox4_Validated(sender As Object, e As EventArgs) Handles ComboBox4.Validated
        ErrorProvider1.SetError(ComboBox4, "")
    End Sub
    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        '''''check if invoice is already generated for selected month, year and godown number
        If bValidateinvoice = True And GrpAddCorrect <> "" Then
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

            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dag = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE from [BILL] where [BILL].bill_date>=Format('" & fDate & "', 'dd/mm/yyyy') and [BILL].bill_date<=Format('" & foDate & "', 'dd/mm/yyyy') and [INVOICE_NO]='" & Trim(TextBox2.Text) & "'", MyConn)
            dsg = New DataSet
            dsg.Clear()
            dag.Fill(dsg, "BILL")

            If dsg.Tables(0).Rows.Count > 0 And GrpAddCorrect <> "C" Then
                Dim errorMsg As String = "Duplicate Invoice Number not allowed..."
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox2, errorMsg)
            End If
            dag.Dispose()
            dsg.Dispose()
            MyConn.Close() ' close connection
        End If

        If DataGridView2.RowCount > 1 And GrpAddCorrect = "A" And cmdClicked = False Then
            get_display_data()
        End If
    End Sub
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidategodown = False
        bValidatetype = False
        bValidateinvoice = False
        bValidatedate = False
    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidategodown = True
        bValidatetype = True
        bValidateinvoice = True
        bValidatedate = True
    End Sub
    Public Function getinvoicesr()

        ''''to get invoice reference number in series
        Try
            Dim INVNO As String
            Dim INVNOTMP As String
            Dim nom As Integer

            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            Dim STR As String = "SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [BILL].SRNO"
            chkrs1.Open("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by [BILL].SRNO", xcon)
            ' chkrs1.MoveLast()
            Do While chkrs1.EOF = False

                ' chkrs1.MovePrevious()
                INVNO = chkrs1.Fields(12).Value.ToString.Substring(12, 3)
                INVNOTMP = chkrs1.Fields(12).Value.ToString.Substring(0, 12)
                chkrs1.MoveNext()
            Loop
            chkrs1.Close()
            xcon.Close()
            nom = Convert.ToInt32(INVNO) + 1
            INVNO = INVNOTMP & String.Format("{0:000}", nom)
            Return INVNO
        Catch ex As Exception
            MsgBox("Exception: Get invoice SR :" & ex.Message)
        End Try
    End Function
    Function get_display_data()

        '''''when invoice number in serial is got for month-year selected get invoice details using godown,party,rent table
        Dim srno As Integer = 0
        Dim numRec As Integer = 0
        If DataGridView2.RowCount > 1 Then
            numRec = Convert.ToInt32(getinvoicesr().ToString.Substring(12, 3)) - 1
            srno = Convert.ToInt32(getinvoiceno()) - 1
        Else
            srno = 1
        End If

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        chkrs.Open("SELECT * FROM GODOWN where [STATUS]='C' and [GROUP]='" & ComboBox3.Text & "' and GODWN_NO='" & ComboBox4.Text & "' order by [GROUP]+GODWN_NO ", xcon)
        Do While chkrs.EOF = False
            numRec = numRec + 1
            srno = srno + 1
            Dim INVOICE_NO As String
            ' Dim FILE_NO As String
            Dim FILE_NOtmp As String
            FILE_NOtmp = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", numRec)
            TextBox13.Text = FILE_NOtmp
            INVOICE_NO = String.Format("{0:0000}", srno)
            FILE_NO = INVOICE_NO.Replace("/", "_")   'DateTimePicker1.Value.Year & "_" & DateTimePicker1.Value.Month & "_" & chkrs.Fields(0).Value & chkrs.Fields(3).Value.replace("/", "_")
            FILE_NO = FILE_NO.Replace(" ", "_")

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
            Dim amt As Double
            Dim CGST_TAXAMT As Double
            Dim SGST_TAXAMT As Double
            Dim CGST_RATE As Double
            Dim SGST_RATE As Double
            Dim gst As Double
            Dim gst_amt As Double
            TENANT_CODE = chkrs.Fields(1).Value
            Label13.Text = ""
            chkrs1.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
            If chkrs1.EOF = False Then
                TENANT_NAME = LTrim(chkrs1.Fields(1).Value)
                Label13.Text = TENANT_NAME

                If IsDBNull(chkrs1.Fields(2).Value) Then
                    TAD1 = ""
                Else
                    If Trim(chkrs1.Fields(2).Value).Equals("") Then
                        TAD1 = ""
                    Else
                        TAD1 = chkrs1.Fields(2).Value.replace("& vbLf & vbLf", "")
                        Label13.Text = Label13.Text + Environment.NewLine + chkrs1.Fields(2).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(3).Value) Then
                    TAD2 = ""
                Else
                    If Trim(chkrs1.Fields(3).Value).Equals("") Then
                        TAD2 = ""
                    Else
                        TAD2 = chkrs1.Fields(3).Value
                        Label13.Text = Label13.Text + Environment.NewLine + chkrs1.Fields(3).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(4).Value) Then
                    TAD3 = ""
                Else
                    If Trim(chkrs1.Fields(4).Value).Equals("") Then
                        TAD3 = ""
                    Else
                        TAD3 = chkrs1.Fields(4).Value
                        Label13.Text = Label13.Text + Environment.NewLine + chkrs1.Fields(4).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(5).Value) Then
                    TCITY = ""
                Else
                    If Trim(chkrs1.Fields(5).Value).Equals("") Then
                        TCITY = ""
                    Else
                        TCITY = chkrs1.Fields(5).Value
                        Label13.Text = Label13.Text + Environment.NewLine + chkrs1.Fields(5).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(17).Value) Then
                    TSTATE = ""
                Else
                    If Trim(chkrs1.Fields(17).Value).Equals("") Then
                        TSTATE = ""
                    Else
                        TSTATE = chkrs1.Fields(17).Value
                        Label13.Text = Label13.Text + Environment.NewLine + chkrs1.Fields(17).Value
                    End If
                End If
                STATE_CODE = "24"

                If IsDBNull(chkrs1.Fields(18).Value) Then
                    TEMAIL = ""
                Else
                    TEMAIL = chkrs1.Fields(18).Value
                    TextBox11.Text = chkrs1.Fields(18).Value
                End If
                If IsDBNull(chkrs1.Fields(19).Value) Then
                    TGST = ""
                Else
                    TGST = chkrs1.Fields(19).Value
                    TextBox12.Text = chkrs1.Fields(19).Value
                End If
            End If

            chkrs1.Close()
            If IsDBNull(chkrs.Fields(37).Value) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = chkrs.Fields(37).Value
            End If


            ComboBox1.Text = DateTimePicker1.Value.Month
            ComboBox2.Text = DateTimePicker1.Value.Year
            chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and P_CODE ='" & TENANT_CODE & "' order by DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
            amt = 0
            If chkrs2.EOF = False Then
                chkrs2.MoveFirst()

                amt = chkrs2.Fields(4).Value
                If IsDBNull(chkrs2.Fields(5).Value) Then
                Else
                    amt = amt + chkrs2.Fields(5).Value
                End If
            End If
            TextBox4.Text = Format(amt, "#####0.00")
            TextBox5.Text = Format(amt, "#####0.00")
            chkrs2.Close()
            Dim ENDDAY As String
            ENDDAY = DateTime.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month).ToString
            If ENDDAY = "31" Then
                ENDDAY = "31st"
            Else
                ENDDAY = ENDDAY + "th"
            End If

            If IsDBNull(chkrs.Fields(37).Value) Or chkrs.Fields(37).Value.Equals("997211") Then
                Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
            Else
                Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
            End If

            chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs.Fields(37).Value & "'", xcon)
            CGST_RATE = 0
            SGST_RATE = 0
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
            gst_amt = gst * amt / 100

            Dim net As Double
            Dim rnd As Integer
            rnd = gst_amt - Math.Round(gst_amt)
            If rnd >= 50 Then
                gst_amt = Math.Round(gst_amt) + 1
            Else
                gst_amt = Math.Round(gst_amt)
            End If

            net = amt + gst_amt
            CGST_TAXAMT = gst_amt / 2


            'CGST_TAXAMT = amt * CGST_RATE / 100
            CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
            'SGST_TAXAMT = amt * SGST_RATE / 100
            SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)

            TextBox6.Text = Format(CGST_TAXAMT, "#####0.00")
            TextBox7.Text = Format(SGST_TAXAMT, "#####0.00")
            TextBox9.Text = Format(CGST_RATE, "#####0.00")
            TextBox10.Text = Format(SGST_RATE, "#####0.00")

            TextBox8.Text = Format(net, "#####0.00")
            If chkrs.EOF = False Then
                chkrs.MoveNext()
            End If
            If chkrs.EOF = True Then
                Exit Do
            End If
        Loop
        chkrs.Close()
        xcon.Close()
    End Function
    Function print_display_data()
        ''''''''generate invoice. 
        Dim objPRNSetup = New clsPrinterSetup
        prnmaxpagelines = objPRNSetup.LinesPerPage
        If objPRNSetup.PageSize = PRNA4Paper Then
            prnleftmargin = 7
        Else
            prnleftmargin = 2
        End If


        Dim array() As String = {"AE", "AA", "AB", "AC"}

        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        Dim srno As Integer = 0
        'xpage = 1
        xpage = Val("2")
        Dim i1 As Double
        Dim numRec As Integer = 0
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        chkrs.Open("SELECT * FROM GODOWN where [STATUS]='C' and [GROUP]='" & ComboBox3.Text & "' and GODWN_NO='" & ComboBox4.Text & "' order by [GROUP]+GODWN_NO ", xcon)
        Do While chkrs.EOF = False
            numRec = numRec + 1
            srno = srno + 1
            Dim INVOICE_NO As String
            ' Dim FILE_NO As String
            Dim FILE_NOtmp As String
            FILE_NOtmp = TextBox13.Text    'DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", numRec)
            INVOICE_NO = TextBox2.Text    'String.Format("{0:0000}", srno)
            FILE_NO = INVOICE_NO.Replace("/", "_")   'DateTimePicker1.Value.Year & "_" & DateTimePicker1.Value.Month & "_" & chkrs.Fields(0).Value & chkrs.Fields(3).Value.replace("/", "_")
            FILE_NO = FILE_NO.Replace(" ", "_")

            If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
            End If
            FileOpen(fnum, Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat", OpenMode.Output)
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
            Dim amt As Double
            Dim CGST_TAXAMT As Double
            Dim SGST_TAXAMT As Double
            Dim CGST_RATE As Double
            Dim SGST_RATE As Double
            Dim gst As Double
            Dim gst_amt As Double
            TENANT_CODE = chkrs.Fields(1).Value
            Label13.Text = ""
            chkrs1.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
            If chkrs1.EOF = False Then
                TENANT_NAME = LTrim(chkrs1.Fields(1).Value)

                If IsDBNull(chkrs1.Fields(2).Value) Then
                    TAD1 = ""
                Else
                    If Trim(chkrs1.Fields(2).Value).Equals("") Then
                        TAD1 = ""
                    Else
                        TAD1 = chkrs1.Fields(2).Value.replace("& vbLf & vbLf", "")
                    End If
                End If
                If IsDBNull(chkrs1.Fields(3).Value) Then
                    TAD2 = ""
                Else
                    If Trim(chkrs1.Fields(3).Value).Equals("") Then
                        TAD2 = ""
                    Else
                        TAD2 = chkrs1.Fields(3).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(4).Value) Then
                    TAD3 = ""
                Else
                    If Trim(chkrs1.Fields(4).Value).Equals("") Then
                        TAD3 = ""
                    Else
                        TAD3 = chkrs1.Fields(4).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(5).Value) Then
                    TCITY = ""
                Else
                    If Trim(chkrs1.Fields(5).Value).Equals("") Then
                        TCITY = ""
                    Else
                        TCITY = chkrs1.Fields(5).Value
                    End If
                End If
                If IsDBNull(chkrs1.Fields(17).Value) Then
                    TSTATE = ""
                Else
                    If Trim(chkrs1.Fields(17).Value).Equals("") Then
                        TSTATE = ""
                    Else
                        TSTATE = chkrs1.Fields(17).Value
                    End If
                End If
                STATE_CODE = "24"

                If IsDBNull(chkrs1.Fields(18).Value) Then
                    TEMAIL = ""
                Else
                    TEMAIL = chkrs1.Fields(18).Value
                End If
                If IsDBNull(chkrs1.Fields(19).Value) Then
                    TGST = ""
                Else
                    TGST = chkrs1.Fields(19).Value
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
            Print(fnum, GetStringToPrint(30, "STATE :" & TSTATE, "S") & Space(15) & "GODOWN NO.  :" & GetStringToPrint(20, chkrs.Fields(0).Value & chkrs.Fields(3).Value, "S") & vbNewLine)
            '  End If
            ' If TGST <> "" Then
            Print(fnum, "GSTIN :" & GetStringToPrint(30, TGST, "S") & Space(8) & "INVOICE NO. :" & GetStringToPrint(35, INVOICE_NO, "S") & vbNewLine)
            '  End If
            '  If TEMAIL <> "" Then
            Print(fnum, "EMAIL ID:" & GetStringToPrint(33, TEMAIL, "S") & Space(3) & "INVOICE DATE:" & GetStringToPrint(20, DateTimePicker1.Value.ToString("dd/MM/yyyy"), "S") & vbNewLine)
            '   End If
            chkrs1.Close()
            Print(fnum, StrDup(30, " ") & vbNewLine)
            Print(fnum, StrDup(30, " ") & "TAX INVOICE FOR SERVICES" & vbNewLine)
            Print(fnum, StrDup(85, "-") & vbNewLine)
            Print(fnum, GetStringToPrint(8, "HSN", "S") & GetStringToPrint(28, "HSN DESCRIPTION", "S") & GetStringToPrint(30, "DESCRIPTION OF SERVICES", "S") & GetStringToPrint(19, "AMOUNT", "N") & vbNewLine)
            Print(fnum, StrDup(85, "-") & vbNewLine)
            If IsDBNull(chkrs.Fields(37).Value) Then
                Print(fnum, GetStringToPrint(7, "", "S"))
            Else
                Print(fnum, GetStringToPrint(7, chkrs.Fields(37).Value, "S"))

            End If
            Print(fnum, GetStringToPrint(28, " Rental Or Leasing Services ", "S"))
            Print(fnum, GetStringToPrint(41, " Rent for property from " & "1st " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S"))
            ComboBox1.Text = DateTimePicker1.Value.Month
            ComboBox2.Text = DateTimePicker1.Value.Year

            amt = TextBox4.Text
            Print(fnum, GetStringToPrint(9, Format(amt, "#####0.00"), "N") & vbNewLine)
            Dim ENDDAY As String
            ENDDAY = DateTime.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month).ToString
            If ENDDAY = "31" Then
                ENDDAY = "31st"
            Else
                ENDDAY = ENDDAY + "th"
            End If
            Print(fnum, GetStringToPrint(7, "", "S"))
            Print(fnum, GetStringToPrint(28, " Involving Own Or Leased ", "S"))
            Print(fnum, GetStringToPrint(35, " to " & ENDDAY & " " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S") & vbNewLine)
            Print(fnum, GetStringToPrint(7, "", "S"))
            If IsDBNull(chkrs.Fields(37).Value) Or chkrs.Fields(37).Value.Equals("997211") Then
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
            Print(fnum, Space(58) & GetStringToPrint(17, "TAXABLE AMOUNT :", "S") & GetStringToPrint(10, Format(amt, "#####0.00"), "N") & vbNewLine)
            Print(fnum, StrDup(85, "-") & vbNewLine)

            'CGST_TAXAMT = amt * CGST_RATE / 100
            'CGST_TAXAMT = Math.Round(CGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
            'SGST_TAXAMT = amt * SGST_RATE / 100
            'SGST_TAXAMT = Math.Round(SGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
            Dim net As Double
            Dim rnd As Integer


            net = TextBox8.Text
            CGST_TAXAMT = TextBox6.Text
            SGST_TAXAMT = TextBox7.Text
            CGST_RATE = TextBox9.Text
            SGST_RATE = TextBox10.Text

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
            ''''''''''''''''''''''''''''''''SEARCH FOR ADVANCE START
            Dim adjamt As Double
            Dim advrec As Integer = 0
            Dim adv_date As Date
            chkrs2.Open("SELECT * FROM RECEIPT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and ADVANCE = TRUE AND ADJ_AMT>0 order by REC_DATE", xcon)
            adjamt = 0
            If chkrs2.EOF = False Then
                chkrs2.MoveFirst()
                advrec = chkrs2.Fields(4).Value
                adv_date = chkrs2.Fields(3).Value
                adjamt = chkrs2.Fields(13).Value - net
                Dim sav As String = "UPDATE [RECEIPT] SET ADJ_AMT=" & adjamt & " WHERE REC_NO=" & chkrs2.Fields(4).Value & " AND year(REC_DATE)='" & Convert.ToDateTime(chkrs2.Fields(3).Value).Year & "'"
                doSQL(sav)
            End If
            chkrs2.Close()
            Print(fnum, Space(40) & GetStringToPrint(45, "For Motilal Hirabhai Estate & Warehouse Ltd.", "S") & vbNewLine)
            If advrec > 0 Then
                Print(fnum, GetStringToPrint(23, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(24, "Received as advance on:", "S") & GetStringToPrint(19, adv_date, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(13, "Receipt No.:", "S") & GetStringToPrint(19, "GST-" + Convert.ToString(advrec), "S") & vbNewLine)
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

            'save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & amt + CGST_TAXAMT + SGST_TAXAMT & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "')"
            Dim save As String
            If GrpAddCorrect = "C" Then
                If advrec > 0 Then
                    save = "UPDATE [BILL] SET P_CODE='" & chkrs.Fields(1).Value & "',BILL_AMOUNT=" & amt & ",CGST_RATE=" & CGST_RATE & ",CGST_AMT=" & CGST_TAXAMT & ",SGST_RATE=" & SGST_RATE & ",SGST_AMT=" & SGST_TAXAMT & ",NET_AMOUNT=" & net & ",HSN='" & chkrs.Fields(37).Value & "',SRNO='" & FILE_NOtmp & "',REC_NO='" & advrec & "',REC_DATE='" & adv_date & "',ADVANCE=TRUE  WHERE INVOICE_NO='" & INVOICE_NO & "' AND [GROUP]='" & chkrs.Fields(0).Value & "' AND GODWN_NO='" & chkrs.Fields(3).Value & "' AND BILL_DATE=#" & DateTimePicker1.Value & "#"
                Else
                    save = "UPDATE [BILL] SET P_CODE='" & chkrs.Fields(1).Value & "',BILL_AMOUNT=" & amt & ",CGST_RATE=" & CGST_RATE & ",CGST_AMT=" & CGST_TAXAMT & ",SGST_RATE=" & SGST_RATE & ",SGST_AMT=" & SGST_TAXAMT & ",NET_AMOUNT=" & net & ",HSN='" & chkrs.Fields(37).Value & "',SRNO='" & FILE_NOtmp & "'  WHERE INVOICE_NO='" & INVOICE_NO & "' AND [GROUP]='" & chkrs.Fields(0).Value & "' AND GODWN_NO='" & chkrs.Fields(3).Value & "' AND BILL_DATE=#" & DateTimePicker1.Value & "#"
                End If
            Else
                If advrec > 0 Then
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO,REC_NO,REC_DATE,ADVANCE) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "','" & advrec & "','" & adv_date & "',TRUE)"
                Else
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "')"
                End If
            End If


            doSQL(save)
            FileClose(fnum)
            pdfpath = Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month)
            strReportFilePath = Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat"

            If chkrs.EOF = False Then
                chkrs.MoveNext()
            End If
            If chkrs.EOF = True Then
                Exit Do
            End If
        Loop
        chkrs.Close()
        xcon.Close()

        'RichTextBox2.LoadFile(strReportFilePath, RichTextBoxStreamType.PlainText)

        ' Catch ex As Exception
        '    MessageBox.Show("Error opening file-sr: " & ex.Message)
        'End Try
    End Function
    Private Sub ComboBox3_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox3.Validating
        Dim errorMsg As String = "Please Select godown type"
        If bValidatetype = True And ComboBox3.Text.Trim.Equals("") Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox3.Select(0, ComboBox3.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(ComboBox3, errorMsg)
        End If
    End Sub

    Private Sub ComboBox3_Validated(sender As Object, e As EventArgs) Handles ComboBox3.Validated
        ErrorProvider1.SetError(ComboBox3, "")
    End Sub

    Private Sub FrmInvSingle_Move(sender As Object, e As EventArgs) Handles Me.Move
        ''''keep the position of form fix on MDI form
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        fillgodowncombo()
    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        cmdClicked = True
        If ValidateChildren() Then
            'doSQL(save)
            print_display_data()    '''''generate invoice

            '''''''create directory if not exist
            If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
            End If
            If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
            End If

            strReportFilePath = Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat"
            CreatePDF(strReportFilePath, FILE_NO)
            MsgBox("Bill is generated at " + Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) + "\ path")
            DataGridView2.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            disablefields()

            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub doSQL(ByVal sql As String)
        '''''method to insrt/update data in database tables
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        Dim objcmd As New OleDb.OleDbCommand
        Try
            objcmd.Connection = MyConn
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sql
            objcmd.ExecuteNonQuery()
            ' MsgBox("Data Inserted successfully in database", vbInformation)
            objcmd.Dispose()
            MyConn.Close()
        Catch ex As Exception
            MsgBox("Exception: Data Insertion in PARTY table in database" & ex.Message)
        End Try
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        ''''''''convert.dat file to pdf file
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
            If ChkLogo.Checked Then
                Dim image As XImage = image.FromFile(Application.StartupPath & "\logo.png")
                graph.DrawImage(image, 0, 0, image.Width, image.Height)
            End If

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

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try
            GrpAddCorrect = "C"
            Label23.Text = "EDIT"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            '  rownum = DataGridView2.CurrentRow.Index
            ChkLogo.Enabled = True
            TextBox12.Focus()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub

    Private Sub DateTimePicker1_Validated(sender As Object, e As EventArgs) Handles DateTimePicker1.Validated
        ErrorProvider1.SetError(DateTimePicker1, "")
    End Sub

    Private Sub DateTimePicker1_Validating(sender As Object, e As CancelEventArgs) Handles DateTimePicker1.Validating
        '''''''check bulk invoices are generated? if not give error to use invoice generate menu option
        If bValidatedate = True And GrpAddCorrect <> "" Then
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dag = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "'", MyConn)
            dsg = New DataSet
            dsg.Clear()
            dag.Fill(dsg, "BILL")
            If dsg.Tables(0).Rows.Count = 0 And GrpAddCorrect <> "C" Then
                Dim errorMsg As String = "Please use Invoice Generate option from menu..."
                e.Cancel = True
                DateTimePicker1.Select()
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(DateTimePicker1, errorMsg)
            End If
            dag.Dispose()
            dsg.Dispose()
            MyConn.Close() ' close connection
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If

            Dim array() As String = {"AE", "AA", "AB", "AC"}

        End If
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ''''''''date picker date always last day of the month
        Dim selectedDate As Date = DateTimePicker1.Value
        Dim DaysInMonth As Integer = Date.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month)
        If (DateTimePicker1.Value.Day <> DaysInMonth) Then
            DateTimePicker1.Value = Convert.ToDateTime(DaysInMonth.ToString + "/" + DateTimePicker1.Value.Month.ToString + "/" + DateTimePicker1.Value.Year.ToString)
        End If
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
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