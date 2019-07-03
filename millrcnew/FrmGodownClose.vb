Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
''' <summary>
''' Tables used - CLGDWN,godown,party
''' In this form we are collecting information about godown he want to close 
''' for some time or suspend for some time
''' </summary>
Public Class FrmGodownClose
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
    Dim indexorder As String = "GODWN_NO"       ''''variable to store sorting order field for datagrid
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
    Public lastdate As Date = DateTime.Today

    Private Sub FrmGodownClose_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        cmdAdd.Enabled = True
        cmdClose.Enabled = True
        cmdUpdate.Enabled = False
        cmdCancel.Enabled = False
        disablefields()
        fillgroupcombo()
        fillgodowncombo()
        ShowData()          ''''''''show data in datagrid view
        LodaDataToTextBox()   ''''''load data to form
        formloaded = True
    End Sub
    Public Function fillgodowncombo()
        '''''''fill godown combo box using godown table with current godowns
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
        ''''fill godown group combo box using group table
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
    Function disablefields()
        ''''disable form fields
        RichTextBox1.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        DateTimePicker1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
    End Function
    Private Sub ShowData()
        ''''show closed /suspended godowns in data grid view using clgdwn table
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE order by [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(4).Visible = False
            DataGridView2.Columns(6).Visible = False

            DataGridView2.Columns(0).HeaderText = "Group"    '''' godown group code - GROUP from clgdwn table
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(3).Width = 71
            DataGridView2.Columns(2).HeaderText = "Godown"   '''' godown number - godwn_no from clgdwn table
            DataGridView2.Columns(7).HeaderText = "Tenant"   '''''tenant name - P_name from party table
            DataGridView2.Columns(3).HeaderText = "Date"     '''''closing date - to_d from clgdwn table
            DataGridView2.Columns(5).HeaderText = "Reason"   '''''reason - reason from cldgwan table
            DataGridView2.Columns(7).Width = 250
            DataGridView2.Columns(5).Width = 100
            Label21.Text = "Total : " & DataGridView2.RowCount - 1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub LodaDataToTextBox()
        '''''load current data from data grid view to form
        Try
            Dim i As Integer
            Label13.Text = ""
            RichTextBox1.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            DateTimePicker1.Text = ""
            If DataGridView2.RowCount > 1 Then
                i = DataGridView2.CurrentRow.Index

                If Not IsDBNull(DataGridView2.Item(0, i).Value) Then
                    ComboBox3.Text = GetValue(DataGridView2.Item(0, i).Value)    '''' godown group code - GROUP from clgdwn table
                End If
                If Not IsDBNull(DataGridView2.Item(2, i).Value) Then
                    ComboBox4.Text = GetValue(DataGridView2.Item(2, i).Value)    '''' godown number - godwn_no from clgdwn table
                End If
                If Not IsDBNull(DataGridView2.Item(3, i).Value) Then
                    DateTimePicker1.Value = GetValue(DataGridView2.Item(3, i).Value)   '''''closing date - to_d from clgdwn table
                End If

                If Not IsDBNull(DataGridView2.Item(7, i).Value) Then
                    TextBox1.Text = GetValue(DataGridView2.Item(1, i).Value)           ''''''tenant code - p_code from cldgwn table
                    Label13.Text = GetValue(DataGridView2.Item(7, i).Value)            ''''''tenant name - p_name from party table
                End If
                If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                    RichTextBox1.Text = GetValue(DataGridView2.Item(5, i).Value)       ''''''reason - reason from clgdwn table
                End If
                If Not IsDBNull(DataGridView2.Item(6, i).Value) Then
                    If GetValue(DataGridView2.Item(6, i).Value).Equals("S") Then       '''''''close/suspended  - close_suspend from clgdwn table
                        RadioButton2.Checked = True
                        RadioButton1.Checked = False
                    Else
                        RadioButton2.Checked = False
                        RadioButton1.Checked = True
                    End If
                Else
                    If GetValue(DataGridView2.Item(6, i).Value).Equals("S") Then
                        RadioButton2.Checked = True
                        RadioButton1.Checked = False
                    Else
                        RadioButton2.Checked = False
                        RadioButton1.Checked = True
                    End If

                End If
                If xcon.State = ConnectionState.Open Then
                Else
                    xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
                End If

                '''''''''select tenant detail from party table for that p_code and append it in label
                chkrs4.Open("SELECT * FROM PARTY WHERE P_CODE='" & GetValue(DataGridView2.Item(1, i).Value) & "'", xcon)
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
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(17).Value
                        End If
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
    Private Sub cmdFirst_Click(sender As Object, e As EventArgs) Handles cmdFirst.Click
        ''''''go to 1st row of the data grid view
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(0).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(0)
        '  rownum = 0
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
        ''''''go to previous row of the data grid view
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
        ''''''go to next row of the data grid view
        Dim intRow As Integer = DataGridView2.CurrentRow.Index
        If intRow < DataGridView2.RowCount - 1 Then
            DataGridView2.CurrentRow.Selected = False
            DataGridView2.Rows(intRow + 1).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(intRow + 1).Cells(0)
            LodaDataToTextBox()
        End If
    End Sub

    Private Sub cmdLast_Click(sender As Object, e As EventArgs) Handles cmdLast.Click
        ''''''go to last row of the data grid view
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(DataGridView2.RowCount - 1).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(0)
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        '''''this event will fire when user write anything in search text box
        '''''It will filter data from gdtrans table and bind with datagrid view
        MyConn = New OleDbConnection(connString)
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
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
        If e.ColumnIndex = 6 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
        Else
            indexorder = "[CLGDWN].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
        End If
        LodaDataToTextBox()
    End Sub
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Try
            Me.Close()   ''' close the form
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyDown
        ''''''''when user select datagrid view row using up, down arrow key and press enter - show that row data in form
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub navigatedisable()
        '''''''disable navigation buttons
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
        TxtSrch.Enabled = False
    End Sub

    Private Sub navigateenable()
        ''''''''enable navigation buttons
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        TxtSrch.Enabled = True
    End Sub

    Private Sub DataGridView2_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyUp
        ''''''''when user select datagrid view row using up, down arrow key and press enter - show that row data in form
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        ''''''''when user double click datagrid row - show that row data in form
        LodaDataToTextBox()
    End Sub
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidategodown = False
        bValidatetype = False

    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidategodown = True
        bValidatetype = True

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
            disablefields()
            navigateenable()
            ShowData()
            LodaDataToTextBox()
            Label23.Text = "VIEW"
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try

    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
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
            ComboBox3.Enabled = True
            RadioButton1.Enabled = True
            RadioButton1.Checked = True
            RadioButton2.Enabled = True
            TextBox1.Text = ""
            ComboBox3.SelectedIndex = ComboBox3.Items.IndexOf("")
            ComboBox3.Text = ""
            ComboBox4.Enabled = True
            ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
            ComboBox4.Text = ""
            ComboBox3.Select()
            Label13.Text = ""
            RichTextBox1.Text = ""
            RichTextBox1.Enabled = True
            DateTimePicker1.Enabled = True
            If IsDBNull(lastdate) Then
                DateTimePicker1.Value = DateTime.Today
                lastdate = DateTimePicker1.Value
            Else
                DateTimePicker1.Value = lastdate
            End If
            '  DateTimePicker1.Value = Date.Today
            navigatedisable()
            cmdAdd.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If GrpAddCorrect = "A" Then
            lastdate = DateTimePicker1.Value     ''''''assign date picker value to lastdate variable
        End If
    End Sub
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs)
        Try
            GrpAddCorrect = "C"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            navigatedisable()
            cmdAdd.Enabled = False
            ComboBox3.Focus()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
        End Try
    End Sub
    Private Sub insertData()
        Dim save, saveold, savenew As String
        Dim stus As String
        If RadioButton1.Checked = True Then
            stus = "D"    ''' for godown closed
        Else
            stus = "S"    ''''for godown suspended
        End If

        If GrpAddCorrect = "C" Then
            ' save = "UPDATE [GODOWN] SET [GROUP]='" & ComboBox1.SelectedValue.ToString & "',P_CODE='" & TextBox10.Text & "',GODWN_NO='" & TextBox2.Text & "',SURVEY='" & TextBox1.Text & "',CENSES='" & TextBox3.Text & "',STATUS='" & stus & "',FROM_D='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "',MONTH_FR='" & DateTimePicker2.Value.Month & "',YEAR_FR='" & DateTimePicker2.Value.Year & "',WIDTH='" & TextBox4.Text & "',LENGTH='" & TextBox5.Text & "',SQ='" & TextBox6.Text & "',MY_FLG='" & RichTextBox1.Text & "',REMARK1='',REMARK2='',REMARK3='',GST='" & ComboBox3.SelectedValue.ToString & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "' AND P_CODE='" & ComboBox2.SelectedValue.ToString & "'"  ' sorry about that
        Else

            save = "INSERT INTO [CLGDWN]([GROUP],P_CODE,GODWN_NO,TO_D,FROM_D,REASON,CLOSE_SUSPEND) VALUES('" & ComboBox3.SelectedValue.ToString & "','" & TextBox1.Text & "','" & ComboBox4.SelectedValue.ToString & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "','" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "','" & RichTextBox1.Text & "','" & stus & "')"

            saveold = "UPDATE [GODOWN] SET [STATUS]='" & stus & "' WHERE [GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox4.Text & "' AND P_CODE='" & TextBox1.Text & "'"  ' sorry about that
        End If
        doSQL(save)      '''''insert data in clgdwn table
        doSQL(saveold)   ''''''update status 'S' or 'D' for closed godown in godown table
        DataGridView2.Update()
        MsgBox("Data Inserted successfully in database", vbInformation)
        '  frmload = True
        ' tabrec = 0
        GrpAddCorrect = ""
        ShowData()
    End Sub
    Private Sub doSQL(ByVal sql As String)
        '''''''method for insert/update data in database tables
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
    Private Sub textdisable()
        ''''''''''disable form elements
        Try
            ComboBox3.Enabled = False
            ComboBox4.Enabled = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            RichTextBox1.Enabled = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Text disable module: " & ex.Message)
        End Try
    End Sub
    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click

        If ValidateChildren() Then
            insertData()    '''''insert / update data in clgdwn table and godown table
            DataGridView2.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
                cmdAdd.Enabled = True
                textdisable()
            Label23.Text = "VIEW"
            GrpAddCorrect = ""
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub KeyUpHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.F1 Then
        End If
    End Sub

    Private Sub KeyDownHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True

    End Sub

    Private Sub FrmGodownClose_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        '''''on F1 key press saw godown details in datagrid view
        If e.KeyCode = Keys.F1 Then
            DataGridView2.Visible = True
            GroupBox5.Visible = True
            GroupBox4.Visible = True
            Me.Width = Me.Width + DataGridView2.Width - 100
            Me.Height = Me.Height + 80
            ShowData()
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        '''''fill godown numbers in godown combo box using godown table when group combo value change
        fillgodowncombo()
        ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
        ComboBox4.Text = ""
        Label13.Text = ""
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        '''''''when user select godown number from godown combo box, select party detail for godown table's p_code from party table and append details in label

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        If ComboBox4.SelectedIndex >= 0 Then
            chkrs1.Open("SELECT * FROM GODOWN WHERE [GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox4.SelectedValue.ToString & "' AND [STATUS]='C'", xcon)
            Do While chkrs1.EOF = False
                Dim PCOD As String = chkrs1.Fields(1).Value
                TextBox1.Text = PCOD
                '''''''''party detail start
                chkrs4.Open("SELECT * FROM PARTY WHERE P_CODE='" & PCOD & "'", xcon)
                Label13.Text = chkrs4.Fields(1).Value
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
                            Label13.Text = Label13.Text + Environment.NewLine + chkrs4.Fields(17).Value
                        End If
                    End If
                End If
                chkrs4.Close()
                ''''''''party detail end
                DateTimePicker2.Value = chkrs1.Fields(11).Value
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs1.Close()
        End If
        xcon.Close()
    End Sub

    Private Sub ComboBox3_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox3.Validating
        Dim errorMsg As String = "Please Select godown type"
        If bValidatetype = True And ComboBox3.Text.Trim.Equals("") And GrpAddCorrect = "A" Then
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
    Private Sub ComboBox4_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox4.Validating
        Dim errorMsg As String = "Please Select godown Number"
        If bValidategodown = True And ComboBox4.Text.Trim.Equals("") And GrpAddCorrect = "A" Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox4.Select(0, ComboBox4.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(ComboBox4, errorMsg)
        End If


    End Sub

    Private Sub ComboBox4_Validated(sender As Object, e As EventArgs) Handles ComboBox4.Validated
        ErrorProvider1.SetError(ComboBox4, "")
    End Sub

    Private Sub FrmGodownClose_Move(sender As Object, e As EventArgs) Handles Me.Move
        '''''to keep the position of for fix on MDI form
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
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