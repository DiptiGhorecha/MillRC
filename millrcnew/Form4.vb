Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Public Class FrmGodown
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
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
    Dim oldName As String
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
    Private Sub FrmGodown_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False
            ' Frame1.Visible = False
            Label13.Width = Me.Width
            Label13.Top = 0
            PictureBox1.Left = 0
            PictureBox1.Top = 0
            cmdAdd.Enabled = True
            cmdClose.Enabled = True
            cmdDelete.Enabled = True
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            textdisable()


            GrpAddCorrect = ""
            fillgroupcombo()
            fillpartycombo()
            fillgstcombo()

            ShowData()
            LodaDataToTextBox()
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
    Public Function fillrentgrid(grp As String, gdn As String, pcd As String)
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            dar = New OleDb.OleDbDataAdapter("SELECT * from [RENT] where [RENT].GROUP='" & grp & "' AND [RENT].GODWN_NO='" & gdn & "' AND [RENT].P_CODE='" & pcd & "' order by DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) desc", MyConn)
            dsr = New DataSet
            If DataGridView1.RowCount > 1 Then
                dsr.Clear()
            End If
            dar.Fill(dsr, "RENT")
            DataGridView1.DataSource = dsr.Tables("RENT")

            dar.Dispose()
            dsr.Dispose()
            MyConn.Close() ' close connection
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                DataGridView1.Columns(i).Visible = False
            Next

            DataGridView1.Columns(4).Visible = True
            DataGridView1.Columns(5).Visible = True
            DataGridView1.Columns(6).Visible = True
            DataGridView1.Columns(8).Visible = True
            DataGridView1.Columns(4).HeaderText = "Rent"
            DataGridView1.Columns(5).HeaderText = "Pe Rent"
            DataGridView1.Columns(6).HeaderText = "From Month"
            DataGridView1.Columns(8).HeaderText = "From Year"
            DataGridView1.Columns(4).Width = 55
            DataGridView1.Columns(5).Width = 90
            DataGridView1.Columns(6).Width = 90
            DataGridView1.Columns(8).Width = 90
            If GrpAddCorrect = "A" Then
                DataGridView1.Rows.Clear()
                '  fillrentgrid(DataGridView2.Item(0, i).Value, DataGridView2.Item(3, i).Value, DataGridView2.Item(1, i).Value)
            End If
            'DataGridView2.Rows(1).Selected = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
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
            ComboBox3.DataSource = dst.Tables("GST")
            ComboBox3.DisplayMember = "GST_DESC"
            ComboBox3.ValueMember = "HSN_NO"
            dat.Dispose()
            dst.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Gst combo fill :" & ex.Message)
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
            For i = 0 To ComboBox1.Items.Count - 1
                '      authors.Add(ComboBox1.Items(i).ToString)
            Next i
            '    ComboBox1.AutoCompleteMode = AutoCompleteMode.Suggest
            'ComboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
            '        ComboBox1.AutoCompleteCustomSource = authors
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
            da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
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

            'DataGridView2.Rows(1).Selected = True
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
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
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
    Private Sub LodaDataToTextBox()
        Try
            Dim i As Integer
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
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            RichTextBox1.Text = ""
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
            If Not IsDBNull(DataGridView2.Item(3, i).Value) Then
                TextBox2.Text = GetValue(DataGridView2.Item(3, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(23, i).Value) Or Not IsDBNull(DataGridView2.Item(24, i).Value) Or Not IsDBNull(DataGridView2.Item(25, i).Value) Or Not IsDBNull(DataGridView2.Item(22, i).Value) Then
                RichTextBox1.Text = GetValue(DataGridView2.Item(22, i).Value) & vbCrLf & GetValue(DataGridView2.Item(23, i).Value) & vbCrLf & GetValue(DataGridView2.Item(24, i).Value) & vbCrLf & GetValue(DataGridView2.Item(25, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(38, i).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(38, i).Value)
                TextBox10.Text = GetValue(DataGridView2.Item(1, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(4, i).Value) Then
                TextBox1.Text = GetValue(DataGridView2.Item(4, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                TextBox3.Text = GetValue(DataGridView2.Item(5, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(18, i).Value) Then
                TextBox4.Text = GetValue(DataGridView2.Item(18, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(19, i).Value) Then
                TextBox5.Text = GetValue(DataGridView2.Item(19, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(20, i).Value) Then
                TextBox6.Text = GetValue(DataGridView2.Item(20, i).Value)
            End If
            If Not IsDBNull(DataGridView2.Item(11, i).Value) Then
                DateTimePicker1.Value = GetValue(DataGridView2.Item(11, i).Value)
            End If

            If Not IsDBNull(DataGridView2.Item(10, i).Value) Then
                If GetValue(DataGridView2.Item(10, i).Value) = "C" Then
                    RadioButton1.Select()
                Else
                    If GetValue(DataGridView2.Item(10, i).Value) = "D" Then
                        RadioButton2.Select()
                    Else
                        RadioButton3.Select()
                    End If
                End If
            End If
            If Not IsDBNull(DataGridView2.Item(14, i).Value) Or Not IsDBNull(DataGridView2.Item(15, i).Value) Then
                DateTimePicker2.Value = New Date(GetValue(DataGridView2.Item(15, i).Value), GetValue(DataGridView2.Item(14, i).Value), 1)
            End If
            If Not IsDBNull(DataGridView2.Item(37, i).Value) Then
                ' ComboBox3.Text = GetValue(DataGridView2.Item(37, i).Value)
                If GetValue(DataGridView2.Item(37, i).Value) = "997212" Then
                    ComboBox3.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                Else
                    ComboBox3.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                End If
            End If
            If Not IsDBNull(DataGridView2.Item(21, i).Value) Then
                TextBox9.Text = GetValue(DataGridView2.Item(21, i).Value)
            End If

            fillrentgrid(GetValue(DataGridView2.Item(0, i).Value), GetValue(DataGridView2.Item(3, i).Value), GetValue(DataGridView2.Item(1, i).Value))
            Label21.Text = "Total : " & DataGridView2.RowCount
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub
    Private Sub textenable()
        Try
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            RichTextBox1.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            DateTimePicker3.Enabled = True
            ' RadioButton1.Enabled = True
            ' RadioButton2.Enabled = True
            '  RadioButton3.Enabled = True
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub textdisable()
        Try
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            TextBox9.Enabled = False
            RichTextBox1.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            DateTimePicker3.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            RadioButton3.Enabled = False
            Exit Sub

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
        da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")
        DataGridView2.DataSource = ds.Tables("GODOWN")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then
            indexorder = "[GODOWN].GROUP"
            GroupBox5.Text = "Search by Group Type"
            '    DataGridView2.Sort(DataGridView2.Columns(0), SortOrder.Descending)
        End If
        If e.ColumnIndex = 3 Then
            indexorder = "[GODOWN].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 38 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView2_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseDoubleClick
        'If e.ColumnIndex = 0 Then
        '    indexorder = "[GODOWN].GROUP"
        '    GroupBox5.Text = "Search by Group Type"
        '    DataGridView2.Sort(DataGridView2.Columns(0), SortOrder.Ascending)
        'End If
        'If e.ColumnIndex = 3 Then
        '    indexorder = "[GODOWN].GODWN_NO"
        '    GroupBox5.Text = "Search by Godown"
        '    DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Ascending)
        'End If
        'If e.ColumnIndex = 38 Then
        '    indexorder = "[PARTY].P_NAME"
        '    GroupBox5.Text = "Search by tenant name"
        '    DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Ascending)
        'End If
        'LodaDataToTextBox()
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
            Label23.Text = "VIEW"
            ErrorProvider1.Clear()
            DataGridView2.Enabled = True
            frmload = True
            tabrec = 0
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            textdisable()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim selectedDate As Date = DateTimePicker1.Value

        If (DateTimePicker1.Value.Day <> 1) Then
            DateTimePicker1.Value = Convert.ToDateTime("1/" + DateTimePicker1.Value.Month.ToString + "/" + DateTimePicker1.Value.Year.ToString)
        End If

    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Dim selectedDate As Date = DateTimePicker2.Value

        If (DateTimePicker2.Value.Day <> 1) Then
            DateTimePicker2.Value = Convert.ToDateTime("1/" + DateTimePicker2.Value.Month.ToString + "/" + DateTimePicker2.Value.Year.ToString)
        End If

    End Sub
    Private Sub dataGridView2_DataBindingComplete(ByVal sender As Object,
    ByVal e As DataGridViewBindingCompleteEventArgs) _
    Handles DataGridView2.DataBindingComplete

        '' Put each of the columns into programmatic sort mode.
        'For Each column As DataGridViewColumn In DataGridView2.Columns
        '    column.SortMode = DataGridViewColumnSortMode.Programmatic
        'Next
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            Label23.Text = "ADD"
            DataGridView2.Enabled = False
            DataGridView1.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            TextBox2.Text = ""
            TextBox1.Enabled = True
            TextBox1.Text = ""
            TextBox2.Enabled = True
            TextBox3.Text = ""
            TextBox3.Enabled = True
            TextBox4.Text = "0"
            TextBox4.Enabled = True
            TextBox5.Text = "0"
            TextBox5.Enabled = True
            TextBox6.Text = "0"
            TextBox6.Enabled = True
            TextBox7.Text = ""
            TextBox7.Enabled = True
            TextBox8.Text = ""
            TextBox8.Enabled = True
            TextBox9.Text = ""
            TextBox9.Enabled = True
            RichTextBox1.Text = ""
            RichTextBox1.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
            ComboBox1.Text = ""
            ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
            ComboBox2.Text = ""
            ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox3.Text = ""
            ComboBox1.Select()
            RadioButton1.Checked = True
            DateTimePicker1.Value = New Date(Now.Year, Now.Month, 1)
            DateTimePicker2.Value = New Date(Now.Year, Now.Month, 1)
            DateTimePicker3.Value = New Date(Now.Year, Now.Month, 1)
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
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

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try
            GrpAddCorrect = "C"
            Label23.Text = "EDIT"
            DataGridView2.Enabled = False
            DataGridView1.Enabled = False
            DateTimePicker2.Enabled = True
            oldName = TextBox1.Text
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            TextBox2.Enabled = False
            TextBox1.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            ComboBox1.Enabled = False
            RichTextBox1.Enabled = True
            ComboBox3.Enabled = True
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
            Label23.Text = "VIEW"
            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub insertData()
        Dim save As String
        Dim stus As String
        Dim gsttype As String
        Dim transaction As OleDbTransaction
        If RadioButton1.Checked Then
            stus = "C"
        Else
            If RadioButton2.Checked Then
                stus = "D"
            Else
                stus = "S"
            End If
        End If

        gsttype = ComboBox3.SelectedValue.ToString
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
            save = "UPDATE [GODOWN] SET [GROUP]='" & ComboBox1.SelectedValue.ToString & "',P_CODE='" & TextBox10.Text & "',GODWN_NO='" & TextBox2.Text & "',SURVEY='" & TextBox1.Text & "',CENSES='" & TextBox3.Text & "',STATUS='" & stus & "',FROM_D='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "',MONTH_FR='" & DateTimePicker2.Value.Month & "',YEAR_FR='" & DateTimePicker2.Value.Year & "',OUTST=" & TextBox9.Text & ",WIDTH='" & TextBox4.Text & "',LENGTH='" & TextBox5.Text & "',SQ='" & TextBox6.Text & "',MY_FLG='" & RichTextBox1.Text & "',REMARK1='',REMARK2='',REMARK3='',GST='" & gsttype & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "' AND P_CODE='" & ComboBox2.SelectedValue.ToString & "'"  ' sorry about that
        Else
            save = "INSERT INTO [GODOWN]([GROUP],P_CODE,GODWN_NO,SURVEY,CENSES,STATUS,FROM_D,MONTH_FR,YEAR_FR,WIDTH,LENGTH,SQ,MY_FLG,GST) VALUES('" & ComboBox1.SelectedValue.ToString & "','" & TextBox10.Text & "','" & TextBox2.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "','" & stus & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & DateTimePicker2.Value.Month & "','" & DateTimePicker2.Value.Year & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & RichTextBox1.Text & "','" & gsttype & "')"
        End If
        objcmd.CommandText = save
        objcmd.ExecuteNonQuery()
        ' doSQL(save)
        If TextBox7.Text.Trim <> "" And TextBox8.Text.Trim = "" Then
            save = "INSERT INTO [RENT](P_CODE,[GROUP],GODWN_NO,RENT,FR_MONTH,FR_YEAR) VALUES('" & TextBox10.Text & "','" & ComboBox1.SelectedValue.ToString & "','" & TextBox2.Text & "','" & TextBox7.Text & "'," & DateTimePicker3.Value.Month & "," & DateTimePicker3.Value.Year & ")"
        Else
            If TextBox7.Text.Trim <> "" And TextBox8.Text.Trim <> "" Then
                save = "INSERT INTO [RENT](P_CODE,[GROUP],GODWN_NO,RENT,PRENT,FR_MONTH,FR_YEAR) VALUES('" & TextBox10.Text & "','" & ComboBox1.SelectedValue.ToString & "','" & TextBox2.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "'," & DateTimePicker3.Value.Month & "," & DateTimePicker3.Value.Year & ")"
            End If
        End If
        objcmd.CommandText = save
        objcmd.ExecuteNonQuery()
        transaction.Commit()
        ' doSQLRent(save) 
        DataGridView2.Update()
        MsgBox("Data Inserted successfully in database", vbInformation)
        '  frmload = True
        ' tabrec = 0
        GrpAddCorrect = ""
        ShowData()
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function


    Private Sub doSQLRent(ByVal sql As String)
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
            '      MsgBox("Data Inserted successfully in database", vbInformation)
            objcmd.Dispose()
        Catch ex As Exception
            MsgBox("Exception: Data Insertion in RENT table in database" & ex.Message)
        End Try
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
            MsgBox("Exception: Data Insertion in PARTY table in database" & ex.Message)
        End Try
    End Sub
    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If bValidategodown = True And GrpAddCorrect <> "" Then
            Dim errorMsg As String = "Please Enter Godown Number"
            If TextBox2.Text.Trim.Equals("") Then
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox2, errorMsg)
            End If

            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dagv = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] where [GODWN_NO]='" & Trim(TextBox2.Text) & "' AND [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND STATUS='C'", MyConn)
            dsgv = New DataSet
            dsgv.Clear()
            dagv.Fill(dsgv, "GODOWN")

            If dsgv.Tables(0).Rows.Count > 0 And GrpAddCorrect <> "C" Then
                errorMsg = "Duplicate Godown Number not allowed..."
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox2, errorMsg)
            End If
            dagv.Dispose()
            dsgv.Dispose()
            MyConn.Close() ' close connection
        End If
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        ErrorProvider1.SetError(TextBox2, "")
    End Sub

    Private Sub ComboBox1_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox1.Validating
        Dim errorMsg As String = "Please Select godown type"
        If bValidatetype = True And ComboBox1.Text.Trim.Equals("") Then
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

    Private Sub ComboBox2_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox2.Validating
        Dim errorMsg As String = "Please Select Tenant"
        If bValidatepname = True And ComboBox2.Text.Trim.Equals("") Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox2.Select(0, ComboBox2.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(ComboBox2, errorMsg)
        End If
    End Sub

    Private Sub ComboBox2_Validated(sender As Object, e As EventArgs) Handles ComboBox2.Validated
        ErrorProvider1.SetError(ComboBox2, "")
    End Sub

    Private Sub ComboBox3_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox3.Validating
        Dim errorMsg As String = "Please Select HSN"
        If bValidateHSN = True And ComboBox3.Text.Trim.Equals("") Then
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
    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        TextBox6.Text = TextBox4.Text * TextBox5.Text
    End Sub
    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        TextBox6.Text = TextBox4.Text * TextBox5.Text
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        DataGridView2.Enabled = False
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        da = New OleDb.OleDbDataAdapter("SELECT * from [BILL] where [P_CODE]='" & TextBox10.Text & "' AND [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "'", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "BILL")

        If ds.Tables(0).Rows.Count > 0 Then
            MsgBox("This data is already used for Invoice... Delete that record first..")
            DataGridView2.Enabled = True
            Exit Sub
        End If
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' close connection


        Try
            Dim kk As Integer = MsgBox("[" & Trim(TextBox1.Text) & "]  Delete Record ?", vbYesNo + vbDefaultButton2)
            If kk = 6 Then  'i assume yes
                MyConn = New OleDbConnection(connString)
                If MyConn.State = ConnectionState.Closed Then
                    MyConn.Open()
                End If
                Dim objcmd As New OleDb.OleDbCommand
                Try
                    objcmd.Connection = MyConn
                    objcmd.CommandType = CommandType.Text
                    objcmd.CommandText = "delete from [GODOWN] where P_CODE='" & ComboBox2.SelectedValue.ToString & "' AND [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "'"
                    objcmd.ExecuteNonQuery()
                    MsgBox("Data deleted successfully from GODOWN table in database", vbInformation)
                    objcmd.Dispose()
                    MyConn.Close()
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    DataGridView2.Enabled = True
                    da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
                    ds = New DataSet
                    ds.Clear()
                    da.Fill(ds)
                    DataGridView2.DataSource = ds.Tables(0).DefaultView
                    DataGridView2.Update()
                    da.Dispose()
                    ds.Dispose()
                    MyConn.Close() ' close connection
                    LodaDataToTextBox()
                Catch ex As Exception
                    MsgBox("Exception: Data Delete module " & ex.Message)
                End Try
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Delete module: " & ex.Message)
        End Try
    End Sub

    Private Sub TxtSrch_TextChanged(sender As Object, e As EventArgs) Handles TxtSrch.TextChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <> -1 Then
            TextBox10.Text = ComboBox2.SelectedValue.ToString
        End If
    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As EventArgs) Handles ComboBox2.GotFocus
        fillpartycombo()
        If GrpAddCorrect = "C" Then
            If Not IsDBNull(DataGridView2.Item(38, DataGridView2.CurrentRow.Index).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(38, DataGridView2.CurrentRow.Index).Value)
                TextBox10.Text = GetValue(DataGridView2.Item(1, DataGridView2.CurrentRow.Index).Value)
            End If
        End If
    End Sub


    Private Sub TextBox7_Validated(sender As Object, e As EventArgs) Handles TextBox7.Validated
        ErrorProvider1.SetError(TextBox7, "")
    End Sub

    Private Sub TextBox7_Validating(sender As Object, e As CancelEventArgs) Handles TextBox7.Validating
        If bValidaterent = True And (GrpAddCorrect <> "" And GrpAddCorrect <> "C") Then
            Dim errorMsg As String = "Please Enter Rent"
            If TextBox7.Text.Trim.Equals("") Then
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                TextBox7.Select(0, TextBox7.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox7, errorMsg)
            End If
        End If
    End Sub

    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown

        If GrpAddCorrect = "C" Then
            If Not IsDBNull(DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value) Then
                ComboBox2.Text = GetValue(DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value)
                ' TextBox10.Text = GetValue(DataGridView2.Item(1, DataGridView2.CurrentRow.Index).Value)
            End If
        Else
            ' fillgroupcombo()
        End If
    End Sub

    Private Sub FrmGodown_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
        If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
        If (Left < 0) Then Left = 0
        If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate
        If ComboBox1.FindString(ComboBox1.Text) < 0 Then
            ComboBox1.Text = ComboBox1.Text.Remove(ComboBox1.Text.Length - 1)
            ComboBox1.SelectionStart = ComboBox1.Text.Length
        End If
    End Sub

    Private Sub ComboBox2_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox2.TextUpdate
        If ComboBox2.FindString(ComboBox2.Text) < 0 Then
            ComboBox2.Text = ComboBox2.Text.Remove(ComboBox2.Text.Length - 1)
            ComboBox2.SelectionStart = ComboBox2.Text.Length
        End If
    End Sub

    Private Sub ComboBox3_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox3.TextUpdate
        If ComboBox3.FindString(ComboBox3.Text) < 0 Then
            ComboBox3.Text = ComboBox3.Text.Remove(ComboBox3.Text.Length - 1)
            ComboBox3.SelectionStart = ComboBox3.Text.Length
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox7_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyUp

        Dim isDigit As Boolean = Char.IsDigit(ChrW(e.KeyValue))
        Dim isKeypadNum As Boolean = e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9
        Dim isBackOrSlashOrPeriod As Boolean = (e.KeyCode = Keys.Decimal Or e.KeyCode = Keys.Oem2 Or e.KeyCode = Keys.Back Or e.KeyCode = Keys.OemPeriod)

        'If Not (isDigit Or isKeypadNum Or isBackOrSlashOrPeriod) Then
        '    '    MessageBox.Show("That's not numeric!")
        '    e.KeyCode = Nothing
        'End If
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
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox4.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox5.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox6.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox9.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class