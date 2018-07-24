Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Public Class FrmGodownReopen
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
    Dim indexorder As String = "GODWN_NO"
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

    Private Sub FrmGodownReopen_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        ShowData()
        LodaDataToTextBox()
        formloaded = True
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
    Function disablefields()
        ' TextBox1.Enabled = False
        RichTextBox1.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        DateTimePicker1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
    End Function
    Private Sub ShowData()
        '  konek() 'open our connection
        Try
            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where [CLGDWN].[CLOSE_SUSPEND]<>'C' order by [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
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

            DataGridView2.Columns(0).HeaderText = "Group"
            DataGridView2.Columns(2).Width = 71
            DataGridView2.Columns(0).Width = 51
            DataGridView2.Columns(3).Width = 71
            DataGridView2.Columns(2).HeaderText = "Godown"
            DataGridView2.Columns(7).HeaderText = "Tenant"
            DataGridView2.Columns(3).HeaderText = "Date"
            DataGridView2.Columns(5).HeaderText = "Reason"
            DataGridView2.Columns(7).Width = 250
            DataGridView2.Columns(5).Width = 100
            Label21.Text = "Total : " & DataGridView2.RowCount - 1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub LodaDataToTextBox()
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
                    ComboBox3.Text = GetValue(DataGridView2.Item(0, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(2, i).Value) Then
                    ComboBox4.Text = GetValue(DataGridView2.Item(2, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(3, i).Value) Then
                    DateTimePicker1.Value = GetValue(DataGridView2.Item(3, i).Value)
                End If

                If Not IsDBNull(DataGridView2.Item(7, i).Value) Then
                    TextBox1.Text = GetValue(DataGridView2.Item(1, i).Value)
                    Label13.Text = GetValue(DataGridView2.Item(7, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(5, i).Value) Then
                    RichTextBox1.Text = GetValue(DataGridView2.Item(5, i).Value)
                End If
                If Not IsDBNull(DataGridView2.Item(6, i).Value) Then
                    If GetValue(DataGridView2.Item(6, i).Value).Equals("S") Then
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

        '  DataGridView1_DoubleClick(DataGridView1, New DataGridViewRowEventArgs(1))
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(0).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(0)
        '  rownum = 0
        LodaDataToTextBox()
    End Sub
    Private Sub cmdPrev_Click(sender As Object, e As EventArgs) Handles cmdPrev.Click
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
        DataGridView2.CurrentRow.Selected = False
        DataGridView2.Rows(DataGridView2.RowCount - 1).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(0)
        '   rownum = DataGridView2.RowCount - 1
        LodaDataToTextBox()
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [CLGDWN].*,[PARTY].P_NAME from [CLGDWN] INNER JOIN [PARTY] on [CLGDWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [CLGDWN].TO_D,[CLGDWN].GROUP,[CLGDWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
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
            Me.Close()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyDown
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
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

    Private Sub DataGridView2_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView2.KeyUp
        If e.KeyCode.Equals(Keys.Enter) Then
            LodaDataToTextBox()
        End If
    End Sub
    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
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
        Try
            GrpAddCorrect = ""
            '''  ErrorProvider1.Clear()
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
        Try
            GrpAddCorrect = "A"
            Label23.Text = "REOPEN"
            DataGridView2.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            DateTimePicker2.Enabled = True
            If IsDBNull(lastdate) Then
                DateTimePicker2.Value = DateTime.Today
                lastdate = DateTimePicker2.Value
            Else
                DateTimePicker2.Value = lastdate
            End If
            '  DateTimePicker1.Value = Date.Today
            navigatedisable()
            cmdAdd.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error Add module: " & ex.Message)
        End Try
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If GrpAddCorrect = "A" Then
            lastdate = DateTimePicker1.Value
        End If
    End Sub
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs)
        Try
            GrpAddCorrect = "C"
            DataGridView2.Enabled = False
            ' cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            navigatedisable()
            cmdAdd.Enabled = False
            '  rownum = DataGridView2.CurrentRow.Index
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
            stus = "D"
        Else
            stus = "S"
        End If

        If GrpAddCorrect = "C" Then
            ' save = "UPDATE [GODOWN] SET [GROUP]='" & ComboBox1.SelectedValue.ToString & "',P_CODE='" & TextBox10.Text & "',GODWN_NO='" & TextBox2.Text & "',SURVEY='" & TextBox1.Text & "',CENSES='" & TextBox3.Text & "',STATUS='" & stus & "',FROM_D='" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "',MONTH_FR='" & DateTimePicker2.Value.Month & "',YEAR_FR='" & DateTimePicker2.Value.Year & "',WIDTH='" & TextBox4.Text & "',LENGTH='" & TextBox5.Text & "',SQ='" & TextBox6.Text & "',MY_FLG='" & RichTextBox1.Text & "',REMARK1='',REMARK2='',REMARK3='',GST='" & ComboBox3.SelectedValue.ToString & "' WHERE [GROUP]='" & ComboBox1.SelectedValue.ToString & "' AND GODWN_NO='" & TextBox2.Text & "' AND P_CODE='" & ComboBox2.SelectedValue.ToString & "'"  ' sorry about that
        Else
            save = "UPDATE[CLGDWN] SET [CLOSE_SUSPEND]='C',REOPEN_DATE='" & Convert.ToDateTime(DateTimePicker2.Value.ToString) & "' WHERE [GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox4.Text & "' AND P_CODE='" & TextBox1.Text & "'"
            saveold = "UPDATE [GODOWN] SET [STATUS]='C' WHERE [GROUP]='" & ComboBox3.SelectedValue.ToString & "' AND GODWN_NO='" & ComboBox4.Text & "' AND P_CODE='" & TextBox1.Text & "'"  ' sorry about that
        End If
        doSQL(save)
        doSQL(saveold)
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
            MsgBox("Exception: Data Insertion Reopen godown " & ex.Message)
        End Try
    End Sub
    Private Sub textdisable()
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
        fillgodowncombo()
        ComboBox4.SelectedIndex = ComboBox4.Items.IndexOf("")
        ComboBox4.Text = ""
        Label13.Text = ""
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
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
            '    Me.ErrorProvider1.SetError(ComboBox3, errorMsg)
        End If
    End Sub

    Private Sub ComboBox3_Validated(sender As Object, e As EventArgs) Handles ComboBox3.Validated
        '    ErrorProvider1.SetError(ComboBox3, "")
    End Sub
    Private Sub ComboBox4_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox4.Validating
        Dim errorMsg As String = "Please Select godown Number"
        If bValidategodown = True And ComboBox4.Text.Trim.Equals("") And GrpAddCorrect = "A" Then
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            ComboBox4.Select(0, ComboBox4.Text.Length)
            ' Set the ErrorProvider error with the text to display. 
            '  Me.ErrorProvider1.SetError(ComboBox4, errorMsg)
        End If


    End Sub

    Private Sub ComboBox4_Validated(sender As Object, e As EventArgs) Handles ComboBox4.Validated
        '  ErrorProvider1.SetError(ComboBox4, "")
    End Sub

    Private Sub FrmGodownReopen_Move(sender As Object, e As EventArgs) Handles Me.Move
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