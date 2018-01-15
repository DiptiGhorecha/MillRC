Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Public Class FrmGodownTransfer
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
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
    Dim dap1 As OleDbDataAdapter
    Dim dsp1 As DataSet
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
    Private bValidatenewtenant As Boolean = True
    Private indexorder As String = "GODWN_NO"
    Private frmload As Boolean = True
    Private tabrec As Integer = 0
    Private colnum As Integer = 0
    Private rownum As Integer = 0
    Dim formloaded As Boolean = False
    Private Sub FrmGodownTransfer_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub FrmGodownTransfer_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False
            cmdAdd.Enabled = True
            cmdClose.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            textdisable()
            GrpAddCorrect = ""
            fillgroupcombo()
            fillgodowncombo()
            fillpartycombo()
            fillpartycombo1()
            fillgstcombo()
            ShowData()
            LodaDataToTextBox()
            formloaded = True
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) order by DATE", MyConn)
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
            DataGridView2.Columns(3).Visible = False
            DataGridView2.Columns(4).Visible = False
            DataGridView2.Columns(5).Visible = False
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(0).Visible = False
            DataGridView2.Columns(7).Visible = False
            DataGridView2.Columns(8).Visible = False
            DataGridView2.Columns(0).Visible = True
            DataGridView2.Columns(1).Visible = True
            DataGridView2.Columns(6).Visible = True
            DataGridView2.Columns(7).Visible = True
            DataGridView2.Columns(8).Visible = True

            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(1).Width = 71
            DataGridView2.Columns(6).Width = 71
            DataGridView2.Columns(7).Width = 325
            DataGridView2.Columns(8).Width = 328
            DataGridView2.Columns(1).HeaderText = "Godown"
            DataGridView2.Columns(6).HeaderText = "Transfer Date"
            DataGridView2.Columns(7).HeaderText = "Old Tenant"
            DataGridView2.Columns(8).HeaderText = "New Tenant"

            'DataGridView2.Rows(1).Selected = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            GrpAddCorrect = ""
            ErrorProvider1.Clear()
            DataGridView2.Enabled = True
            frmload = True
            tabrec = 0
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            textdisable()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub
    Private Sub LodaDataToTextBox()
        Try
            Dim i As Integer
            '  TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            ComboBox1.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
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
            If Not IsDBNull(DataGridView2.Item(7, i).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(7, i).Value)
                ' ComboBox2.SelectedIndex = ComboBox2.FindStringExact(ComboBox2.Text)
            End If
            If Not IsDBNull(DataGridView2.Item(8, i).Value) Then
                ComboBox3.Text = GetValue(DataGridView2.Item(8, i).Value)
            End If

            If Not IsDBNull(DataGridView2.Item(1, i).Value) Then
                ComboBox5.Text = GetValue(DataGridView2.Item(1, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(2, i).Value) Then
                TextBox2.Text = GetValue(DataGridView2.Item(2, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                TextBox3.Text = GetValue(DataGridView2.Item(5, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(6, i).Value) Then
                DateTimePicker1.Value = GetValue(DataGridView2.Item(6, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(4, i).Value) Then
                TextBox3.Text = GetValue(DataGridView2.Item(4, i).Value)
            End If


            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            'chkrs.Open("SELECT * FROM GODOWN where [GROUP]='" & DataGridView2.Item(0, i).Value & "' AND [GODWN_NO]='" & DataGridView2.Item(1, i).Value & "' AND [STATUS]='C' order by [GROUP]+GODWN_NO ", xcon)
            chkrs.Open("SELECT * FROM GODOWN where [GROUP]='" & DataGridView2.Item(0, i).Value & "' AND [GODWN_NO]='" & DataGridView2.Item(1, i).Value & "' order by [GROUP]+GODWN_NO ", xcon)
            Do While chkrs.EOF = False
                If Not IsDBNull(chkrs.Fields(4).Value) Then
                    TextBox4.Text = GetValue(chkrs.Fields(4).Value)
                End If
                If Not IsDBNull(chkrs.Fields(5).Value) Then
                    TextBox5.Text = GetValue(chkrs.Fields(5).Value)
                End If
                If Not IsDBNull(chkrs.Fields(18).Value) Then
                    TextBox6.Text = GetValue(chkrs.Fields(18).Value)
                End If
                If Not IsDBNull(chkrs.Fields(19).Value) Then
                    TextBox7.Text = GetValue(chkrs.Fields(19).Value)
                End If
                If Not IsDBNull(chkrs.Fields(20).Value) Then
                    TextBox8.Text = GetValue(chkrs.Fields(20).Value)
                End If
                If Not IsDBNull(chkrs.Fields(37).Value) Then
                    ' ComboBox4.Text = GetValue(chkrs.Fields(37).Value)
                    If GetValue(chkrs.Fields(37).Value) = "997212" Then
                        ComboBox4.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                    Else
                        ComboBox4.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                    End If
                End If
                If chkrs.EOF = False Then
                    chkrs.MoveNext()
                End If
                If chkrs.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs.Close()
            xcon.Close()

            Label21.Text = "Total : " & DataGridView2.RowCount
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            TextBox2.Text = ""
            ' TextBox1.Enabled = True
            '  TextBox1.Text = ""
            ' TextBox2.Enabled = True
            TextBox3.Text = ""
            ' TextBox3.Enabled = True
            TextBox4.Text = "0"
            TextBox4.Enabled = True
            TextBox5.Text = "0"
            TextBox5.Enabled = True
            TextBox6.Text = "0"
            TextBox6.Enabled = True
            TextBox7.Text = ""
            TextBox7.Enabled = True
            TextBox8.Text = ""
            TextBox1.Text = "0"
            TextBox1.Enabled = True
            TextBox9.Text = "0"
            TextBox9.Enabled = True
            TextBox8.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = False
            ComboBox3.Enabled = True
            ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
            ComboBox1.Text = ""
            ComboBox5.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox5.Text = ""
            ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
            ComboBox2.Text = ""
            ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox3.Text = ""
            ComboBox4.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox4.Text = ""
            ComboBox1.Select()
            DateTimePicker1.Value = New Date(Now.Year, Now.Month, 1)
            navigatedisable()
            cmdAdd.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
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
        da = New OleDb.OleDbDataAdapter("SELECT [GDTRANS].*,P2.P_NAME AS PNAME1,P1.P_NAME AS PNAME2 from (([GDTRANS] INNER JOIN [PARTY] AS P2 on [GDTRANS].OP_CODE=P2.P_CODE) INNER JOIN [PARTY] AS P1 ON [GDTRANS].NP_CODE=P1.P_CODE) where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY date", MyConn)
        ds = New DataSet
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
            '    DataGridView2.Sort(DataGridView2.Columns(0), SortOrder.Descending)
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "GODWN_NO"
            GroupBox5.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 6 Then
            indexorder = "[DATE]"
            GroupBox5.Text = "Search by Date"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
        If e.ColumnIndex = 7 Then
            indexorder = "P2.P_NAME"
            GroupBox5.Text = "Search by Old tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
        If e.ColumnIndex = 8 Then
            indexorder = "P1.P_NAME"
            GroupBox5.Text = "Search by New tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
        LodaDataToTextBox()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <> -1 Then
            TextBox2.Text = ComboBox2.SelectedValue.ToString
        End If
    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As EventArgs) Handles ComboBox2.GotFocus
        'fillpartycombo()
        'If GrpAddCorrect = "C" Then
        '    If Not IsDBNull(DataGridView2.Item(38, DataGridView2.CurrentRow.Index).Value) Then
        '        ComboBox2.Text = GetValue(DataGridView2.Item(7, DataGridView2.CurrentRow.Index).Value)
        '        TextBox2.Text = GetValue(DataGridView2.Item(2, DataGridView2.CurrentRow.Index).Value)
        '    End If
        'End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex <> -1 Then
            TextBox3.Text = ComboBox3.SelectedValue.ToString
        End If
    End Sub

    Private Sub ComboBox3_GotFocus(sender As Object, e As EventArgs) Handles ComboBox3.GotFocus
        fillpartycombo1()
        ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
        ComboBox3.Text = ""
        TextBox3.Text = ""
        If GrpAddCorrect = "C" Then
            If Not IsDBNull(DataGridView2.Item(8, DataGridView2.CurrentRow.Index).Value) Then
                ComboBox3.Text = GetValue(DataGridView2.Item(8, DataGridView2.CurrentRow.Index).Value)
                TextBox3.Text = GetValue(DataGridView2.Item(4, DataGridView2.CurrentRow.Index).Value)
            End If
        End If
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
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
    Public Function fillgstcombo()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dat = New OleDb.OleDbDataAdapter("SELECT * from [GST] Order by [GST].GST_DESC", MyConn)
            dst = New DataSet
            dst.Clear()
            dat.Fill(dst, "GST")
            ComboBox4.DataSource = dst.Tables("GST")
            ComboBox4.DisplayMember = "GST_DESC"
            ComboBox4.ValueMember = "HSN_NO"
            dat.Dispose()
            dst.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Gst combo fill :" & ex.Message)
        End Try
    End Function
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
    Public Function fillpartycombo1()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dap1 = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] Order by [PARTY].P_NAME", MyConn)
            dsp1 = New DataSet
            dsp1.Clear()
            dap1.Fill(dsp1, "PARTY")
            ComboBox3.DataSource = dsp1.Tables("PARTY")
            ComboBox3.DisplayMember = "P_NAME"
            ComboBox3.ValueMember = "P_CODE"
            dap1.Dispose()
            dsp1.Dispose()
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
    Private Sub textenable()
        Try
            'TextBox1.Enabled = True
            '  TextBox2.Enabled = True
            ' TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            ComboBox1.Enabled = True
            '  ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True
            DateTimePicker1.Enabled = True
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Text enable module: " & ex.Message)
        End Try
    End Sub

    Private Sub textdisable()
        Try
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
            ComboBox5.Enabled = False
            DateTimePicker1.Enabled = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Text disable module: " & ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        fillgodowncombo()
        ComboBox5.SelectedIndex = ComboBox5.Items.IndexOf("")
        ComboBox5.Text = ""
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox8.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
        ComboBox4.Text = ""
        'Label13.Text = ""
    End Sub

    Private Sub ComboBox5_KeyUp(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyUp
        'Dim index As Integer
        'Dim actual As String
        'Dim found As String

        '' Do nothing for some keys such as navigation keys.
        'If ((e.KeyCode = Keys.Back) Or
        '    (e.KeyCode = Keys.Left) Or
        '    (e.KeyCode = Keys.Right) Or
        '    (e.KeyCode = Keys.Up) Or
        '    (e.KeyCode = Keys.Delete) Or
        '    (e.KeyCode = Keys.Down) Or
        '    (e.KeyCode = Keys.PageUp) Or
        '    (e.KeyCode = Keys.PageDown) Or
        '    (e.KeyCode = Keys.Home) Or
        '    (e.KeyCode = Keys.End)) Then

        '    Return
        'End If

        '' Store the actual text that has been typed.
        'actual = Me.ComboBox5.Text

        '' Find the first match for the typed value.
        'index = Me.ComboBox5.FindString(actual)

        '' Get the text of the first match.
        'If (index > -1) Then
        '    found = Me.ComboBox5.Items(index).ToString()

        '    ' Select this item from the list.
        '    Me.ComboBox5.SelectedIndex = index

        '    ' Select the portion of the text that was automatically
        '    ' added so that additional typing will replace it.
        '    Me.ComboBox5.SelectionStart = actual.Length
        '    Me.ComboBox5.SelectionLength = found.Length
        'End If
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        Dim PCOD As String
        If ComboBox5.SelectedIndex >= 0 Then
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If

            'chkrs1.Open("SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.SelectedValue.ToString & "' AND [STATUS]='C'", xcon)
            chkrs1.Open("SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.SelectedValue.ToString & "'", xcon)

            Do While chkrs1.EOF = False
                PCOD = chkrs1.Fields(1).Value
                If Not IsDBNull(chkrs1.Fields(4).Value) Then
                    TextBox4.Text = GetValue(chkrs1.Fields(4).Value)
                End If
                If Not IsDBNull(chkrs1.Fields(5).Value) Then
                    TextBox5.Text = GetValue(chkrs1.Fields(5).Value)
                End If
                If Not IsDBNull(chkrs1.Fields(18).Value) Then
                    TextBox6.Text = GetValue(chkrs1.Fields(18).Value)
                End If
                If Not IsDBNull(chkrs1.Fields(19).Value) Then
                    TextBox7.Text = GetValue(chkrs1.Fields(19).Value)
                End If
                If Not IsDBNull(chkrs1.Fields(20).Value) Then
                    TextBox8.Text = GetValue(chkrs1.Fields(20).Value)
                End If
                If Not IsDBNull(chkrs1.Fields(37).Value) Then
                    'ComboBox4.Text = GetValue(chkrs1.Fields(37).Value)
                    If GetValue(chkrs1.Fields(37).Value) = "997212" Then
                        ComboBox4.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                    Else
                        ComboBox4.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                    End If
                End If
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
    Private Sub insertData()
        Dim save, saveold, savenew As String
        Dim stus As String

        If GrpAddCorrect = "C" Then
            ' save = "UPDATE [GODOWN] SET [GROUP]='" & ComboBox1.SelectedValue.ToString & "',P_CODE='" & TextBox10.Text & "',GODWN_NO='" & TextBox2.Text & "',SURVEY='" & TextBox1.Text & "',CENSES='" & TextBox3.Text & "',STATUS='" & stus & "',FROM_D='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "',MONTH_FR='" & DateTimePicker2.Value.Month & "',YEAR_FR='" & DateTimePicker2.Value.Year & "',WIDTH='" & TextBox4.Text & "',LENGTH='" & TextBox5.Text & "',SQ='" & TextBox6.Text & "',MY_FLG='" & RichTextBox1.Text & "',REMARK1='',REMARK2='',REMARK3='',GST='" & ComboBox3.SelectedValue.ToString & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "' AND P_CODE='" & ComboBox2.SelectedValue.ToString & "'"  ' sorry about that
        Else
            save = "INSERT INTO [GDTRANS]([GROUP],GODWN_NO,OP_CODE,OP_NAME,NP_CODE,NP_NAME,[DATE]) VALUES('" & ComboBox1.SelectedValue.ToString & "','" & ComboBox5.SelectedValue.ToString & "','" & TextBox2.Text & "',' ','" & TextBox3.Text & "',' ','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "')"
            saveold = "UPDATE [GODOWN] SET [STATUS]='D',TO_D='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.Text & "' AND P_CODE='" & TextBox2.Text & "'"  ' sorry about that
            savenew = "INSERT INTO [GODOWN]([GROUP],P_CODE,GODWN_NO,SURVEY,CENSES,STATUS,FROM_D,MONTH_FR,YEAR_FR,WIDTH,LENGTH,SQ,MY_FLG,GST) VALUES('" & ComboBox1.SelectedValue.ToString & "','" & TextBox3.Text & "','" & ComboBox5.SelectedValue.ToString & "','" & TextBox4.Text & "','" & TextBox5.Text & "','C','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & DateTimePicker1.Value.Month & "','" & DateTimePicker1.Value.Year & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "',' ','" & ComboBox4.SelectedValue.ToString & "')"
        End If
        doSQL(save)
        doSQL(saveold)
        doSQL(savenew)
        If TextBox1.Text.Trim <> "" And TextBox9.Text.Trim = "" Then
            save = "INSERT INTO [RENT](P_CODE,[GROUP],GODWN_NO,RENT,FR_MONTH,FR_YEAR) VALUES('" & TextBox3.Text & "','" & ComboBox1.SelectedValue.ToString & "','" & ComboBox5.SelectedValue.ToString & "','" & TextBox1.Text & "'," & DateTimePicker1.Value.Month & "," & DateTimePicker1.Value.Year & ")"
        Else
            If TextBox1.Text.Trim <> "" And TextBox9.Text.Trim <> "" Then
                save = "INSERT INTO [RENT](P_CODE,[GROUP],GODWN_NO,RENT,PRENT,FR_MONTH,FR_YEAR) VALUES('" & TextBox3.Text & "','" & ComboBox1.SelectedValue.ToString & "','" & ComboBox5.SelectedValue.ToString & "','" & TextBox1.Text & "','" & TextBox9.Text & "'," & DateTimePicker1.Value.Month & "," & DateTimePicker1.Value.Year & ")"
            End If
        End If
        doSQL(save)
        DataGridView2.Update()
        MsgBox("Data Inserted successfully in database", vbInformation)
        '  frmload = True
        ' tabrec = 0
        GrpAddCorrect = ""
        ShowData()
    End Sub
    Private Sub doSQL(ByVal sql As String)
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
        Catch ex As Exception
            MsgBox("Exception: Data Insertion Transfer of godown " & ex.Message)
        End Try
    End Sub
    Public Function checkDue(ptrcode As String, ctrDate As Date) As Boolean
        Dim dueamt As Boolean = True
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim invno As String

        'chkrs1.Open("SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.SelectedValue.ToString & "' AND [STATUS]='C'", xcon)
        chkrs3.Open("SELECT * FROM BILL WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox5.SelectedValue.ToString & "' AND P_CODE='" & ptrcode & "' and REC_NO IS NULL AND REC_DATE IS NULL", xcon)
        Do While chkrs3.EOF = False
            invno = chkrs3.Fields(0).Value
            dueamt = False

            Exit Do
        Loop
        chkrs3.Close()
        xcon.Close()

        If dueamt = False Then
            MsgBox("Please clear due....")

        End If
        Return dueamt
    End Function
    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If ValidateChildren() And checkDue(TextBox2.Text, DateTimePicker1.Value) Then
            insertData()
            DataGridView2.Enabled = True
            If GrpAddCorrect = "C" Then
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                textdisable()
            Else
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                textdisable()
            End If
            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub

    End Sub


    Private Sub ComboBox1_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox1.Validating
        Dim errorMsg As String = "Please Select godown type"
        If bValidatetype = True And ComboBox1.Text.Trim.Equals("") And GrpAddCorrect = "A" Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox1.Select(0, ComboBox1.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(ComboBox1, errorMsg)
        End If
    End Sub

    Private Sub ComboBox1_Validated(sender As Object, e As EventArgs) Handles ComboBox1.Validated
        ErrorProvider1.SetError(ComboBox1, "")
    End Sub
    Private Sub ComboBox5_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox5.Validating
        Dim errorMsg As String = "Please Select godown Number"
        If bValidategodown = True And ComboBox5.Text.Trim.Equals("") And GrpAddCorrect = "A" Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox5.Select(0, ComboBox5.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(ComboBox5, errorMsg)
        End If


    End Sub

    Private Sub ComboBox5_Validated(sender As Object, e As EventArgs) Handles ComboBox5.Validated
        ErrorProvider1.SetError(ComboBox5, "")
    End Sub
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidategodown = False
        bValidatetype = False

    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidategodown = True
        bValidatetype = True

    End Sub
    Private Sub ComboBox3_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox3.TextUpdate
        If ComboBox3.FindString(ComboBox3.Text) < 0 Then
            ComboBox3.Text = ComboBox3.Text.Remove(ComboBox3.Text.Length - 1)
            ComboBox3.SelectionStart = ComboBox3.Text.Length

        End If
    End Sub
    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate
        If ComboBox1.FindString(ComboBox1.Text) < 0 Then
            ComboBox1.Text = ComboBox1.Text.Remove(ComboBox1.Text.Length - 1)
            ComboBox1.SelectionStart = ComboBox1.Text.Length
        End If
    End Sub
    Private Sub ComboBox5_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox5.TextUpdate
        If ComboBox5.FindString(ComboBox5.Text) < 0 Then
            ComboBox5.Text = ComboBox5.Text.Remove(ComboBox5.Text.Length - 1)
            ComboBox5.SelectionStart = ComboBox5.Text.Length
        End If
    End Sub
    Private Sub ComboBox4_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox4.TextUpdate
        If ComboBox4.FindString(ComboBox4.Text) < 0 Then
            ComboBox4.Text = ComboBox4.Text.Remove(ComboBox4.Text.Length - 1)
            ComboBox4.SelectionStart = ComboBox4.Text.Length
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox1.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox9.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox6.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox7.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox8.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub

    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        rownum = DataGridView2.CurrentRow.Index
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyDown
        rownum = DataGridView2.CurrentRow.Index
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView2_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyUp
        rownum = DataGridView2.CurrentRow.Index
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        rownum = DataGridView2.CurrentRow.Index
        LodaDataToTextBox()
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub ComboBox3_Validated(sender As Object, e As EventArgs) Handles ComboBox3.Validated
        ErrorProvider1.SetError(ComboBox3, "")
    End Sub

    Private Sub ComboBox3_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox3.Validating
        If bValidatenewtenant = True And GrpAddCorrect <> "" Then


            If ComboBox3.Text.Equals(ComboBox2.Text) And GrpAddCorrect <> "C" Then
                Dim errorMsg As String = "Old Tenant and New Tenant should be not same..."
                e.Cancel = True
                ComboBox3.Select(0, ComboBox3.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(ComboBox3, errorMsg)
            End If
        End If
    End Sub
    Private Sub ComboBox2_Validated(sender As Object, e As EventArgs) Handles ComboBox2.Validated
        ErrorProvider1.SetError(ComboBox2, "")
    End Sub

    Private Sub ComboBox2_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox2.Validating
        If bValidatenewtenant = True And GrpAddCorrect <> "" Then
            If ComboBox2.Text.Equals(ComboBox3.Text) And GrpAddCorrect <> "C" Then
                Dim errorMsg As String = "Old Tenant and New Tenant should be not same..."
                e.Cancel = True
                ComboBox2.Select(0, ComboBox2.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(ComboBox2, errorMsg)
            End If
        End If
    End Sub

    Private Sub TxtSrch_TextChanged(sender As Object, e As EventArgs) Handles TxtSrch.TextChanged

    End Sub
End Class