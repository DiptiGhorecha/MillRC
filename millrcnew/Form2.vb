Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
''' <summary>
''' tables used - group,godown
''' this module is used to add / edit / delete/ search godown group (godown type) 
''' </summary>
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
    Dim strReportFilePath As String   ''''''variable used to store report file name with path
    Dim GrpAddCorrect As String       ''''''variable used to store crud status - C for Edit, A for Add & '' for View
    Dim blnTranStart As Boolean
    Dim formloaded As Boolean = False    '''' variable used to store form load event status
    Dim oldName As String
    Dim ok As Boolean
    Private bValidatepcode As Boolean = True
    Private bValidatepname As Boolean = True
    Private indexorder As String = "G_CODE"    ''''variable to store sorting order field for datagrid



    Private Sub FrmGodownType_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            '''''''Setting form position top left corner of mdi
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False

            '''''''enabled form components
            DataGridView1.Enabled = True
            cmdAdd.Enabled = True
            cmdClose.Enabled = True
            cmdDelete.Enabled = True
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            textdisable()                       ''''diable text boxes
            GrpAddCorrect = ""
            ShowData()                          '''''load data in datagrid from group table
            LodaDataToTextBox()                 '''''load data to form text box
            formloaded = True
            If muser.Equals("super") Then       '''''' staff can not add,edit or delete a godown type
                cmdAdd.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub
    Private Sub textenable()
        '''''function to enable textbox
        Try
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub textdisable()
        '''''function to disable textbox
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
            GrpAddCorrect = "A"              '''''''store crud status variable 
            DataGridView1.Enabled = False    '''''''DISABLED datagrid when user is adding a record 
            cmdUpdate.Enabled = True         '''''''enabled buttons UPDATE and CANCEL 
            cmdCancel.Enabled = True
            textenable()                     '''''''enabled all form text boxes 
            TextBox1.Text = ""               '''''''celared textboxes
            TextBox2.Text = ""
            TextBox2.Enabled = True
            TextBox2.Select()
            navigatedisable()                '''''''disabled  navigation buttons 
            cmdAdd.Enabled = False           ''''''disabled ADD,EDIT,DELETE button
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Label22.Text = "ADD"             ''''''updated status text with ADD to let the user know which crud operation is going on 
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub
    Private Sub navigatedisable()
        ''''''''DISABLED navigation buttons
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
        TxtSrch.Enabled = False
    End Sub

    Private Sub navigateenable()
        '''''''''enable navigation buttons
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        TxtSrch.Enabled = True
    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Try
            GrpAddCorrect = "C"                     '''''''store crud status variable 
            Label22.Text = "EDIT"                   ''''''updated status text with EDIT to let the user know which crud operation is going on 
            DataGridView1.Enabled = False
            oldName = TextBox1.Text
            cmdUpdate.Enabled = True                '''''''enabled buttons UPDATE and CANCEL 
            cmdCancel.Enabled = True
            TextBox2.Enabled = False
            TextBox1.Enabled = True
            TextBox1.Select()
            navigatedisable()                       '''''''disabled  navigation buttons 
            cmdAdd.Enabled = False                  ''''''disabled ADD,EDIT,DELETE button
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        DataGridView1.Enabled = False                ''''''''disabled datagrid
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If

        ''''''''check the selected group for deletion in godown table, if any godown is created for this group type then prompt the user to delete the godown first
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
            Dim kk As Integer = MsgBox("[" & Trim(TextBox1.Text) & "]  Delete Record ?", vbYesNo + vbDefaultButton2)    '''''VB yes/no messagebox to get confirmation from ghe user  
            If kk = 6 Then  ''''''if user click yes
                MyConn = New OleDbConnection(connString)
                If MyConn.State = ConnectionState.Closed Then
                    MyConn.Open()
                End If
                Dim objcmd As New OleDb.OleDbCommand
                Try
                    '''''''''''''''delete current record from group table
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
                    '''''''''''''''delete current record from group table

                    DataGridView1.Enabled = True    ''''''''''''enabled and update datagrid so uer can play with other godown type data
                    da = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] order by " & indexorder, MyConn)
                    ds = New DataSet
                    ds.Clear()
                    da.Fill(ds, "GROUP")
                    DataGridView1.DataSource = ds.Tables("GROUP")
                    DataGridView1.Update()
                    da.Dispose()
                    ds.Dispose()
                    MyConn.Close() ' close connection
                    LodaDataToTextBox()               ''''''''''''''load selected datagrid record to text box
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
        '''''''''''load data from group table to datagrid
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            da = New OleDb.OleDbDataAdapter("SELECT * from [GROUP] order by " & indexorder, MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "GROUP")
            DataGridView1.DataSource = ds.Tables("GROUP")

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection

            '''''''''''''showing only two columns Group code and group description and change it's width
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
        LodaDataToTextBox()   ''''load data to textboxes from the row where user double clicked
    End Sub
    Private Sub LodaDataToTextBox()
        '''''''''''''''''load selected datagrid row's data to text boxes
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            If Not IsDBNull(DataGridView1.Item(0, i).Value) Then
                TextBox2.Text = DataGridView1.Item(0, i).Value    ''''''''g_code from group table
            End If
            If Not IsDBNull(DataGridView1.Item(1, i).Value) Then
                TextBox1.Text = DataGridView1.Item(1, i).Value     '''''''''''g_name from group table
            End If

            Label2.Text = "Total : " & DataGridView1.RowCount        '''''number of records in group table
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Try
            Me.Close()    ''''''''''close the form
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            '''''''''''''replace the text ADD/EDIT with VIEW , DISABLED update,cancel button and enabled ADD,EDIT,DELETE buttons 


            Label22.Text = "VIEW"
            DataGridView1.Enabled = True
            GrpAddCorrect = ""
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            textdisable()      '''''''''''''''disabled textboxes
            navigateenable()   '''''''''''''''enabled navigation buttons
            ShowData()         '''''''''''''''update datagrid 
            LodaDataToTextBox()   '''''''''''load selected datagrid view row's data to textboxes
            Me.ErrorProvider1.Clear()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub
    Private Sub insertData()
        Dim save As String     '''''''''variable to store query string accoarding to ADD or Edit button pressed
        If GrpAddCorrect = "C" Then
            save = "UPDATE [GROUP] SET G_CODE='" & TextBox2.Text & "',G_NAME='" & TextBox1.Text & "' WHERE G_CODE='" & TextBox2.Text & "'" ' sorry about that
        Else
            save = "INSERT INTO [GROUP](G_CODE,G_NAME) VALUES('" & TextBox2.Text & "','" & TextBox1.Text & "')"
        End If
        doSQL(save)    ''''''''update group table
        ShowData()     '''''''show updated data from group table to datagrid
    End Sub
    Private Sub doSQL(ByVal sql As String)
        ''''''method to insert/update data in database table
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
            insertData()                     '''''''insert/update data in group table
            DataGridView1.Enabled = True     '''''''enabled datagrid

            ''''''''' disabled UPDATE,CANCEL buttons and enabled ADD,EDIT,DELETE buttons.
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            textdisable()               '''''''''''disabled textboxes

            Label22.Text = "VIEW"       ''''''''' replaced crud status label with text VIEW
            navigateenable()            ''''''''' enabled navigation buttons
            LodaDataToTextBox()         ''''''''' load data from selected datagrid row to text boxes
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
        ''''''''set selection pointer to 1st row of datagrid and load that record data to text boxes
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(0).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(0)
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        ''''''''set selection pointer to previous row of datagrid and load that record data to text boxes
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow > 0 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow - 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow - 1).Cells(0)
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        ''''''''set selection pointer to next row of datagrid and load that record data to text boxes
        Dim intRow As Integer = DataGridView1.CurrentRow.Index
        If intRow < DataGridView1.RowCount - 1 Then
            DataGridView1.CurrentRow.Selected = False
            DataGridView1.Rows(intRow + 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(intRow + 1).Cells(0)
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        ''''''''set selection pointer to last row of datagrid and load that record data to text boxes
        DataGridView1.CurrentRow.Selected = False
        DataGridView1.Rows(DataGridView1.RowCount - 1).Selected = True
        DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(0)
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        '''''''''search datagrid for the text user type in search text box
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

    Private Sub DataGridView1_SizeChanged(sender As Object, e As EventArgs) Handles DataGridView1.SizeChanged
        ''''''''''''''''update record count label when datagrid is updated
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
        ''''''''load the datagrid row data to textboxes when user select record by pressing enter key
        If e.KeyCode.Equals(Keys.Enter) Then
            e.SuppressKeyPress = True
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ''If e.KeyCode.Equals(Keys.Enter) Then

        ''    LodaDataToTextBox()
        ''End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick

        ''''''''set index order for searching and search textbox label accoarding to datagrid column user clicked 
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
        ''''''''set index order for searching and search textbox label accoarding to datagrid column user clicked 
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
        ''''''''''''''''code to make group form's position fix
        If formloaded Then
            If (Me.Right > Parent.ClientSize.Width) Then Me.Left = Parent.ClientSize.Width - Me.Width
            If (Me.Bottom > Parent.ClientSize.Height) Then Me.Top = Parent.ClientSize.Height - Me.Height
            If (Me.Left < 0) Then Me.Left = 0
            If (Me.Top < 0) Then Me.Top = 0
            If (Me.Top < 87) Then Me.Top = 87
            Me.Refresh()
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case ((m.WParam.ToInt64() And &HFFFF) And &HFFF0)

            Case &HF060 ' The user chose to close the form.
                Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        End Select

        MyBase.WndProc(m)
    End Sub

End Class