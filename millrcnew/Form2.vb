Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Public Class FrmGodownType
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
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim strReportFilePath As String
    Dim GrpAddCorrect As String
    Dim blnTranStart As Boolean
    Dim formloaded As Boolean = False
    Dim oldName As String
    Dim ok As Boolean
    Private bValidatepcode As Boolean = True
    Private bValidatepname As Boolean = True
    Private indexorder As String = "G_CODE"



    Private Sub FrmGodownType_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False
            ' Label4.Width = Me.Width
            '  Label4.Top = 0
            '  PictureBox1.Left = 0
            '  PictureBox1.Top = 0
            ' PictureBox2.Left = Me.Width - 50
            '  PictureBox2.Top = 10
            ' '  PictureBox3.Left = Me.Width - 100
            '  PictureBox3.Top = 10

            ' Frame1.Visible = False
            DataGridView1.Enabled = True
            cmdAdd.Enabled = True
            cmdClose.Enabled = True
            cmdDelete.Enabled = True
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            textdisable()
            GrpAddCorrect = ""
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
    Private Sub textenable()
        Try
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub textdisable()
        Try
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            DataGridView1.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox2.Enabled = True
            TextBox2.Select()
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

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try

            GrpAddCorrect = "C"
            DataGridView1.Enabled = False
            'Datprimaryrs.Recordset.Edit
            oldName = TextBox1.Text
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            TextBox2.Enabled = False
            TextBox1.Enabled = True
            TextBox1.Select()
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        DataGridView1.Enabled = False
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        da = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] where [GROUP]='" & Trim(TextBox2.Text) & "'", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")

        If ds.Tables(0).Rows.Count > 0 Then
            MsgBox("This data is already used in Godown Master.. Delete that record first..")
            DataGridView1.Enabled = True
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
                    objcmd.CommandText = "delete from [GROUP] where G_CODE='" & Trim(TextBox2.Text) & "'"
                    objcmd.ExecuteNonQuery()
                    MsgBox("Data deleted successfully from GROUP table in database", vbInformation)
                    objcmd.Dispose()
                    MyConn.Close()
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    DataGridView1.Enabled = True
                    da = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] order by " & indexorder, MyConn)
                    ds = New DataSet
                    ds.Clear()
                    da.Fill(ds, "GROUP")
                    DataGridView1.DataSource = ds.Tables("GROUP")
                    DataGridView1.Update()
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
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] order by " & indexorder, MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "GROUP")
            DataGridView1.DataSource = ds.Tables("GROUP")

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection
            DataGridView1.Columns(0).HeaderText = "Godown Type"
            DataGridView1.Columns(0).Width = 98
            DataGridView1.Columns(1).Width = 313
            DataGridView1.Columns(1).HeaderText = "Description"
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        LodaDataToTextBox()

    End Sub
    Private Sub LodaDataToTextBox()

        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            If Not IsDBNull(DataGridView1.Item(0, i).Value) Then
                TextBox2.Text = DataGridView1.Item(0, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(1, i).Value) Then
                TextBox1.Text = DataGridView1.Item(1, i).Value
            End If

            Label2.Text = "Total : " & DataGridView1.RowCount
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

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            DataGridView1.Enabled = True
            GrpAddCorrect = ""
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            textdisable()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Me.ErrorProvider1.Clear()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub
    Private Sub insertData()
        Dim save As String
        If GrpAddCorrect = "C" Then
            save = "UPDATE [GROUP] SET G_CODE='" & TextBox2.Text & "',G_NAME='" & TextBox1.Text & "' WHERE G_CODE='" & TextBox2.Text & "'" ' sorry about that
        Else
            save = "INSERT INTO [GROUP](G_CODE,G_NAME) VALUES('" & TextBox2.Text & "','" & TextBox1.Text & "')"
        End If
        doSQL(save)
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
            MsgBox("Data Inserted successfully in database", vbInformation)
            objcmd.Dispose()
        Catch ex As Exception
            MsgBox("Exception: Data Insertion in GROUP table in database" & ex.Message)
        End Try
    End Sub
    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If ValidateChildren() Then
            insertData()
            DataGridView1.Enabled = True
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
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Dim errorMsg As String = "Please Enter Godown Description"
        If bValidatepname = True And TextBox1.Text.Trim.Equals("") Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            TextBox1.Select(0, TextBox1.Text.Length)

            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(TextBox1, errorMsg)
        End If
    End Sub

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        ErrorProvider1.SetError(TextBox1, "")
    End Sub
    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If bValidatepcode = True And GrpAddCorrect <> "" Then
            Dim errorMsg As String = "Please Enter Godown Type"
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
            da = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] where [G_CODE]='" & Trim(TextBox2.Text) & "'", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "GROUP")

            If ds.Tables(0).Rows.Count > 0 And GrpAddCorrect <> "C" Then
                errorMsg = "Duplicate Godown Type not allowed..."
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox2, errorMsg)
            End If
            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection


        End If
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        ErrorProvider1.SetError(TextBox2, "")
    End Sub

    Private Sub cmdFirst_Click(sender As Object, e As EventArgs) Handles cmdFirst.Click
        '  DataGridView1_DoubleClick(DataGridView1, New DataGridViewRowEventArgs(1))
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(0).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(0)
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow > 0 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow - 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow - 1).Cells(0)
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow < DataGridView1.RowCount - 1 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow + 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow + 1).Cells(0)
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(DataGridView1.RowCount - 1).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(0)
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        ' End If
        da = New OleDb.OleDbDataAdapter("SELECT * FROM [GROUP] where " & indexorder & " Like '%" & TxtSrch.Text & "%'", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GROUP")
        DataGridView1.DataSource = ds.Tables("GROUP")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        ' If DataGridView1.RowCount > 0 Then
        'Label2.Text = DataGridView1.CurrentRow.ToString & " / " & DataGridView1.RowCount
        '  End If
    End Sub

    Private Sub DataGridView1_SizeChanged(sender As Object, e As EventArgs) Handles DataGridView1.SizeChanged
        If DataGridView1.RowCount > 0 Then
            Label2.Text = DataGridView1.CurrentRow.ToString & " / " & DataGridView1.RowCount
        End If
    End Sub
    Private Sub DataGridView1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        Dim dv As DataView = New DataView(ds.Tables(0))

        dv.Sort = "G_NAME"
        dv.RowFilter = "G_NAME like '" & e.KeyChar & "%'"
        If dv.Count > 0 Then
            Dim x As String = dv(0)(0).ToString
            Dim bs As New BindingSource
            bs.DataSource = ds.Tables(0)
            DataGridView1.BindingContext(bs).Position = bs.Find("G_NAME", x)
            DataGridView1.CurrentCell = DataGridView1(0, bs.Position)
        End If
    End Sub

    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidatepcode = False
        bValidatepname = False
    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidatepcode = True
        bValidatepname = True
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick


        If e.ColumnIndex = 0 Then
            indexorder = "G_CODE"
            GroupBox1.Text = "Search by godown type"
            DataGridView1.Sort(DataGridView1.Columns(0), SortOrder.Ascending)
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "G_NAME"
            GroupBox1.Text = "Search by description"
            DataGridView1.Sort(DataGridView1.Columns(1), SortOrder.Ascending)
        End If
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseDoubleClick
        If e.ColumnIndex = 0 Then
            indexorder = "G_CODE"
            GroupBox1.Text = "Search by godown type"
            DataGridView1.Sort(DataGridView1.Columns(0), SortOrder.Ascending)
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "G_NAME"
            GroupBox1.Text = "Search by description"
            DataGridView1.Sort(DataGridView1.Columns(1), SortOrder.Ascending)
        End If
        LodaDataToTextBox()
    End Sub

    Private Sub FrmGodownType_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Me.Right > Parent.ClientSize.Width) Then Me.Left = Parent.ClientSize.Width - Me.Width
            If (Me.Bottom > Parent.ClientSize.Height) Then Me.Top = Parent.ClientSize.Height - Me.Height
            If (Me.Left < 0) Then Me.Left = 0
            If (Me.Top < 0) Then Me.Top = 0
            If (Me.Top < 87) Then Me.Top = 87
            Me.Refresh()
        End If
    End Sub

    Private Sub FrmGodownType_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        'If Me.WindowState = FormWindowState.Minimized Or Me.WindowState = FormWindowState.Normal Then
        '    If formloaded Then
        '        If (Me.Right > Parent.ClientSize.Width) Then Me.Left = Parent.ClientSize.Width - Me.Width
        '        If (Me.Bottom > Parent.ClientSize.Height) Then Me.Top = Parent.ClientSize.Height - Me.Height
        '        If (Me.Left < 0) Then Me.Left = 0
        '        If (Me.Top < 0) Then Me.Top = 0
        '        If (Me.Top < 90) Then Me.Top = 90
        '    End If
        'End If


    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
    End Sub
End Class