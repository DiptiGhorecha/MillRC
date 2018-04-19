Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Public Class FrmAdvance
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
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
    Dim oldDate As Date
    Dim hsn As String
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

    Private Sub FrmAdvance_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False
            ' Frame1.Visible = False

            cmdAdd.Enabled = True
            cmdClose.Enabled = True
            cmdDelete.Enabled = True
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            textdisable()
            GrpAddCorrect = ""
            fillgroupcombo()
            fillgodowncombo()
            fillpartycombo()
            ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
            ComboBox5.SelectedIndex = ComboBox5.Items.IndexOf("")
            ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
            ShowData()
            If (DataGridView2.RowCount > 0) Then
                LodaDataToTextBox()
            End If

            formloaded = True
            If muser.Equals("super") Then
                cmdAdd.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
            End If
            '  TextBox1.Text = chkrs1.Fields(1).Value
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        fillgodowncombo()
        ComboBox5.SelectedIndex = ComboBox5.Items.IndexOf("")
        ComboBox5.Text = ""
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""

    End Sub
    Public Function fillgodowncombo()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dap = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] WHERE [GROUP]='" & ComboBox1.Text & "' and [STATUS]='C' Order by GODWN_NO", MyConn)
            dsp = New DataSet
            dsp.Clear()
            dap.Fill(dsp, "GODOWN")
            ComboBox5.DataSource = dsp.Tables("GODOWN")
            ComboBox5.DisplayMember = "GODWN_NO"
            ComboBox5.ValueMember = "GODWN_NO"
            dap.Dispose()
            dsp.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Godown combo fill :" & ex.Message)
        End Try
    End Function
    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        Dim PCOD As String
        If ComboBox5.SelectedIndex >= 0 Then
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If

            'chkrs1.Open("SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.SelectedValue.ToString & "' AND [STATUS]='C'", xcon)
            Dim STR As String = "SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox1.Text & "' AND GODWN_NO='" & ComboBox5.Text & "' and STATUS='C'"
            chkrs1.Open(Str, xcon)

            Do While chkrs1.EOF = False
                PCOD = chkrs1.Fields(1).Value
                TextBox10.Text = PCOD
                hsn = chkrs1.Fields(37).Value
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs1.Close()
            xcon.Close()
            If GrpAddCorrect = "A" And ComboBox5.SelectedIndex >= 0 Then
                fillpartyaddcombo(PCOD)
            End If
        End If
    End Sub

    Public Function fillpartyaddcombo(pcod As String)
        Try
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            chkrs1.Open("SELECT * from [PARTY] where p_code='" & pcod & "' Order by [PARTY].P_NAME", xcon)
            Dim pname As String
            Do While chkrs1.EOF = False
                pname = chkrs1.Fields(1).Value
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs1.Close()
            xcon.Close()
            ComboBox2.Text = pname
        Catch ex As Exception
            MessageBox.Show("Party combo fill :" & ex.Message)
        End Try
    End Function
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim selectedDate As Date = DateTimePicker1.Value
        DateTimePicker1.Value = Convert.ToDateTime(DateTimePicker1.Value.Day.ToString + "/" + DateTimePicker1.Value.Month.ToString + "/" + DateTimePicker1.Value.Year.ToString)
    End Sub
    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            ComboBox5.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
            ComboBox1.Text = ""
            ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
            ComboBox2.Text = ""
            TextBox1.Text = ""
            TextBox10.Text = ""
            ComboBox1.Select()
            DateTimePicker1.Value = Date.Now   'New Date(Now.Year, Now.Month, Date.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month))
            DateTimePicker2.Value = Date.Now   'New Date(Now.Year, Now.Month, Date.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month))
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False

            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub

    Private Sub textenable()
        Try
            ComboBox5.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            TextBox1.Enabled = True
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try
            GrpAddCorrect = "C"
            DataGridView2.Enabled = False
            oldDate = DateTimePicker1.Value.ToString
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            TextBox1.Enabled = True
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            ComboBox5.Enabled = False
            ComboBox1.Enabled = False
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            rownum = DataGridView2.CurrentRow.Index
            ComboBox2.Focus()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If ValidateChildren() Then
            insertData()
            DataGridView2.Enabled = True
            If GrpAddCorrect = "C" Then
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                textdisable()
            Else
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                textdisable()
            End If
            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub insertData()
        Try
            Dim save As String
            Dim transaction As OleDbTransaction
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
                save = "UPDATE [ADVANCES] SET ADVANCE_TILL_DATE='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "', REC_DATE='" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "', REC_NO='" & TextBox1.Text & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.Text & "' AND P_CODE='" & TextBox10.Text & "'"
            Else
                save = "INSERT INTO [ADVANCES]([GROUP],P_CODE,GODWN_NO,ADVANCE_TILL_DATE,REC_DATE,REC_NO) VALUES('" & ComboBox1.SelectedValue.ToString & "','" & TextBox10.Text & "','" & ComboBox5.Text & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "','" & TextBox1.Text & "')"
            End If
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()
            '''''''''''''''''bill update
            Dim invdt As String
            invdt = "30/06/2017"
            Dim lesser, greater As DateTime
            Dim amt As Integer
            Dim CGST_TAXAMT As Double
            Dim SGST_TAXAMT As Double
            Dim CGST_RATE As Double
            Dim gst_amt As Double
            Dim SGST_RATE As Double
            Dim gst As Double

            lesser = Convert.ToDateTime(DateTimePicker1.Value.ToString)
            greater = Convert.ToDateTime(invdt)
            If lesser <= greater Then


                If xcon.State = ConnectionState.Open Then
                Else
                    xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                End If
                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' and GODWN_NO='" & ComboBox5.SelectedValue.ToString & "' and P_CODE ='" & TextBox10.Text & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                amt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()

                    amt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amt = amt + chkrs2.Fields(5).Value
                    End If
                End If
                chkrs2.Close()
                chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & hsn & "'", xcon)
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
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)

                While lesser <= greater
                    Console.WriteLine(lesser.Month)
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO) VALUES('" & "" & "','" & ComboBox1.SelectedValue.ToString & "','" & ComboBox5.SelectedValue.ToString & "','" & TextBox10.Text & "','" & Convert.ToDateTime(lesser) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & hsn & "','" & "" & "')"
                    lesser = lesser.AddMonths(1)
                    objcmd.CommandText = save
                    objcmd.ExecuteNonQuery()
                End While
                xcon.Close()
            End If
            If TextBox1.Text <> "" Then
                If GrpAddCorrect = "C" Then
                save = "UPDATE [BILL] SET ADVANCE= False, REC_DATE=Null, REC_NO=Null  WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.Text & "' AND P_CODE='" & TextBox10.Text & "' AND BILL_DATE<= Format('" & Convert.ToDateTime(oldDate) & "','dd/mm/yyyy')"
                objcmd.CommandText = save
                objcmd.ExecuteNonQuery()
            End If

            save = "UPDATE [BILL] SET ADVANCE= True , REC_DATE ='" & DateTimePicker1.Value.ToShortDateString & "',REC_NO='" & TextBox1.Text & "'  WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.Text & "' AND P_CODE='" & TextBox10.Text & "' AND BILL_DATE<= Format('" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','dd/mm/yyyy')"
            objcmd.CommandText = save
                objcmd.ExecuteNonQuery()
            End If
            transaction.Commit()
            DataGridView2.Update()
            MsgBox("Data Inserted successfully in database", vbInformation)
            GrpAddCorrect = ""
            ShowData()
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub
    Private Sub cmdFirst_Click(sender As Object, e As EventArgs) Handles cmdFirst.Click

        '  DataGridView1_DoubleClick(DataGridView1, New DataGridViewRowEventArgs(1))
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(0).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(0)
        rownum = 0
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        Dim intRow As Integer = DataGridView2.CurrentRow.Index
        If intRow > 0 Then
            DataGridView2.CurrentRow.Selected = False
            DataGridView2.Rows(intRow - 1).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(intRow - 1).Cells(0)
            rownum = intRow - 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        Dim intRow As Integer = DataGridView2.CurrentRow.Index
        If intRow < DataGridView2.RowCount - 1 Then
            DataGridView2.CurrentRow.Selected = False
            DataGridView2.Rows(intRow + 1).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(intRow + 1).Cells(0)
            rownum = intRow + 1
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(DataGridView2.RowCount - 1).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(0)
        rownum = DataGridView2.RowCount - 1
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [ADVANCES].*,[PARTY].P_NAME from [ADVANCES] INNER JOIN [PARTY] on [ADVANCES].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [ADVANCES].GROUP+[ADVANCES].GODWN_NO", MyConn)
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then
            indexorder = "[GROUP]"
            GroupBox5.Text = "Search by Group Type"
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "GODWN_NO"
            GroupBox5.Text = "Search by Godown"
        End If
        If e.ColumnIndex = 4 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
        End If
        If e.ColumnIndex = 3 Then
            indexorder = "Search by Date"
            GroupBox5.Text = ""
        End If
        LodaDataToTextBox()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'If ComboBox2.SelectedIndex <> -1 Then
        '    TextBox10.Text = ComboBox2.SelectedValue.ToString
        'End If
    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As EventArgs) Handles ComboBox2.GotFocus
        fillpartycombo()
        If GrpAddCorrect = "C" Then
            If Not IsDBNull(DataGridView2.Item(2, DataGridView2.CurrentRow.Index).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(6, DataGridView2.CurrentRow.Index).Value)
                '   TextBox10.Text = GetValue(DataGridView2.Item(2, DataGridView2.CurrentRow.Index).Value)
            End If
        End If
    End Sub
    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown

        If GrpAddCorrect = "C" Then
            If Not IsDBNull(DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value)
            End If
        Else

        End If
    End Sub
    Private Sub LodaDataToTextBox()
        Try
            Dim i As Integer

            TextBox10.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox5.Text = ""
            TextBox1.text=""
            DataGridView2.ClearSelection()
            DataGridView2.Rows(rownum).Selected = True
            DataGridView2.FirstDisplayedScrollingRowIndex = rownum
            DataGridView2.CurrentCell = DataGridView2.Rows(rownum).Cells(0)
            If frmload = True Then
                i = 0
                frmload = False
            Else
                i = DataGridView2.CurrentRow.Index
            End If

            If Not IsDBNull(DataGridView2.Item(0, i).Value) Then
                ComboBox1.Text = GetValue(DataGridView2.Item(0, i).Value)
                ComboBox1.SelectedIndex = ComboBox1.FindStringExact(ComboBox1.Text)
            End If
            If Not IsDBNull(DataGridView2.Item(1, i).Value) Then
                ComboBox5.Text = GetValue(DataGridView2.Item(1, i).Value)
                ComboBox5.SelectedIndex = ComboBox5.FindStringExact(ComboBox5.Text)
            End If
            If Not IsDBNull(DataGridView2.Item(6, i).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(6, i).Value)
                TextBox10.Text = GetValue(DataGridView2.Item(2, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                TextBox1.Text = GetValue(DataGridView2.Item(5, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(4, i).Value) Then
                DateTimePicker2.Value = GetValue(DataGridView2.Item(4, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(3, i).Value) Then
                DateTimePicker1.Value = GetValue(DataGridView2.Item(3, i).Value)
            End If


            Label21.Text = "Total : " & DataGridView2.RowCount
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Try
            Me.Close()

            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Public Function fillpartycombo()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dap = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] Order by [PARTY].P_NAME", MyConn)
            dsp = New DataSet
            dsp.Clear()
            dap.Fill(dsp, "PARTY")
            ComboBox2.DataSource = dsp.Tables("PARTY")
            ComboBox2.DisplayMember = "P_NAME"
            ComboBox2.ValueMember = "P_CODE"
            dap.Dispose()
            dsp.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Party combo fill :" & ex.Message)
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
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [ADVANCES].*,[PARTY].P_NAME from [ADVANCES] INNER JOIN [PARTY] on [ADVANCES].P_CODE=[PARTY].P_CODE order by [ADVANCES].GROUP+[ADVANCES].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection
            DataGridView2.Columns(0).Visible = True
            DataGridView2.Columns(1).Visible = True
            DataGridView2.Columns(2).Visible = False
            DataGridView2.Columns(3).Visible = True
            DataGridView2.Columns(4).Visible = True
            DataGridView2.Columns(5).Visible = True
            DataGridView2.Columns(6).Visible = True
            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(3).Width = 71
            DataGridView2.Columns(1).Width = 51
            DataGridView2.Columns(1).HeaderText = "Godown"
            DataGridView2.Columns(6).HeaderText = "Tenant"
            DataGridView2.Columns(3).HeaderText = "Advance till date"
            DataGridView2.Columns(4).HeaderText = "Receipt Date"
            DataGridView2.Columns(3).HeaderText = "Receipt No."
            DataGridView2.Columns(4).Width = 71
            DataGridView2.Columns(5).Width = 57
            DataGridView2.Columns(6).Width = 139
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidategodown = False
        bValidatetype = False
        bValidatepname = False
        bValidateHSN = False
        bValidaterent = False
    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidategodown = True
        bValidatetype = True
        bValidatepname = True
        bValidateHSN = True
        bValidaterent = True
    End Sub
    Private Sub textdisable()
        Try
            ComboBox5.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            TextBox1.Enabled = False
            DateTimePicker2.Enabled = False
            DateTimePicker1.Enabled = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmAdvance_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            GrpAddCorrect = ""
            DataGridView2.Enabled = True
            frmload = True
            tabrec = 0
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            textdisable()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Dim selectedDate As Date = DateTimePicker2.Value

        DateTimePicker2.Value = Convert.ToDateTime(DateTimePicker2.Value.Day.ToString + "/" + DateTimePicker2.Value.Month.ToString + "/" + DateTimePicker2.Value.Year.ToString)
        ' End If
    End Sub
End Class