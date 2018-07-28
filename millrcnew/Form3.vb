Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Class FrmTenant
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
    Dim oldName As String
    Dim ok As Boolean
    Private bValidatepcode As Boolean = True
    Private bValidatepname As Boolean = True
    Private BVALIDATEEMAIL As Boolean = True
    Dim formloaded As Boolean = False
    Private indexorder As String = "P_CODE"
    Private Sub FrmTenant_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Me.MaximizeBox = False
            ' Frame1.Visible = False
            ' Label4.Width = Me.Width
            ' Label4.Top = 0
            '  PictureBox1.Left = 0
            '  PictureBox1.Top = 0
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
            'TextBox1.Text = chkrs1.Fields(1).Value
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
            da = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] order by " & indexorder, MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "PARTY")
            DataGridView1.DataSource = ds.Tables("PARTY")

            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection
            DataGridView1.Columns(0).HeaderText = "Tenant Code"
            DataGridView1.Columns(1).Width = 311
            DataGridView1.Columns(1).HeaderText = "Tenant Name"
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).Visible = False
            DataGridView1.Columns(16).Visible = False
            DataGridView1.Columns(17).Visible = False
            DataGridView1.Columns(19).Visible = False
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
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            '  TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            If DataGridView1.RowCount >= 1 Then
                i = DataGridView1.CurrentRow.Index
            If Not IsDBNull(DataGridView1.Item(0, i).Value) Then
                TextBox2.Text = DataGridView1.Item(0, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(1, i).Value) Then
                TextBox1.Text = DataGridView1.Item(1, i).Value
            End If
            RichTextBox1.Text = ""
            If Not IsDBNull(DataGridView1.Item(2, i).Value) Then
                If IsDBNull(DataGridView1.Item(2, i).Value) Then

                Else
                    If Trim(DataGridView1.Item(2, i).Value).Equals("") Then

                    Else

                        RichTextBox1.Text = RichTextBox1.Text + Replace(DataGridView1.Item(2, i).Value, vbLf, "")
                    End If
                End If
                If IsDBNull(DataGridView1.Item(3, i).Value) Then

                Else
                    If Trim(DataGridView1.Item(3, i).Value).Equals("") Then

                    Else

                        RichTextBox1.Text = RichTextBox1.Text + Environment.NewLine + DataGridView1.Item(3, i).Value
                    End If
                End If

                If IsDBNull(DataGridView1.Item(4, i).Value) Then

            Else
                If Trim(DataGridView1.Item(4, i).Value).Equals("") Then

                Else

                    RichTextBox1.Text = RichTextBox1.Text + Environment.NewLine + DataGridView1.Item(4, i).Value
                End If
            End If




                '  RichTextBox1.Text = DataGridView1.Item(2, i).Value & vbCrLf & DataGridView1.Item(3, i).Value & vbCrLf & DataGridView1.Item(4, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(5, i).Value) Then
                TextBox3.Text = DataGridView1.Item(5, i).Value
            End If
            RichTextBox3.Text = ""
            If Not IsDBNull(DataGridView1.Item(6, i).Value) Then
                RichTextBox3.Text = DataGridView1.Item(6, i).Value & vbCrLf & DataGridView1.Item(7, i).Value & vbCrLf & DataGridView1.Item(8, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(9, i).Value) Then
                TextBox5.Text = DataGridView1.Item(9, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(10, i).Value) Then
                TextBox6.Text = DataGridView1.Item(10, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(11, i).Value) Then
                TextBox4.Text = DataGridView1.Item(11, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(13, i).Value) Then
                TextBox7.Text = DataGridView1.Item(13, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(17, i).Value) Then
                TextBox8.Text = DataGridView1.Item(17, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(18, i).Value) Then
                TextBox9.Text = DataGridView1.Item(18, i).Value
            End If
            If Not IsDBNull(DataGridView1.Item(19, i).Value) Then
                TextBox10.Text = DataGridView1.Item(19, i).Value
            End If
                Label6.Text = "Total : " & DataGridView1.RowCount
            End If
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
            '  TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            RichTextBox1.Enabled = True
            RichTextBox3.Enabled = True

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
            TextBox10.Enabled = False
            RichTextBox1.Enabled = False
            RichTextBox3.Enabled = False
            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
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
        da = New OleDb.OleDbDataAdapter("SELECT * FROM [PARTY] where " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY " & indexorder, MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "PARTY")
        DataGridView1.DataSource = ds.Tables("PARTY")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection

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
            Label22.Text = "VIEW"
            Me.ErrorProvider1.Clear()
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
    Public Function getpcode()
        Dim Sql As String = " Select Case Max(Val(P_CODE)) As Expr1 From PARTY;"
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        da = New OleDb.OleDbDataAdapter("Select Max(Val(P_CODE)) As Expr1 From PARTY", MyConn)
        ds = New DataSet

        ds.Clear()
        da.Fill(ds, "PARTY")
        Dim srno As Integer = Convert.ToInt32(ds.Tables(0).Rows(0)(0)) + 1
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' close connection
        Return srno
    End Function
    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            GrpAddCorrect = "A"
            Label22.Text = "ADD"
            DataGridView1.Enabled = False
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            textenable()
            TextBox2.Text = getpcode()
            TextBox1.Text = ""
            TextBox2.Enabled = True
            TextBox2.Select()
            TextBox3.Text = ""
            TextBox3.Enabled = True
            TextBox4.Text = ""
            TextBox4.Enabled = True
            TextBox5.Text = ""
            TextBox5.Enabled = True
            TextBox6.Text = ""
            TextBox6.Enabled = True
            TextBox7.Text = ""
            TextBox7.Enabled = True
            ' TextBox8.Text = ""
            ' TextBox8.Enabled = True
            TextBox9.Text = ""
            TextBox9.Enabled = True
            TextBox10.Text = ""
            TextBox10.Enabled = True
            RichTextBox1.Text = ""
            RichTextBox1.Enabled = True
            RichTextBox3.Text = ""
            RichTextBox3.Enabled = True
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
            Label22.Text = "EDIT"
            'Datprimaryrs.Recordset.Edit
            DataGridView1.Enabled = False
            oldName = TextBox1.Text
            cmdUpdate.Enabled = True
            cmdCancel.Enabled = True
            TextBox2.Enabled = False
            TextBox1.Enabled = True
            TextBox1.Select()
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            '  TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            RichTextBox1.Enabled = True
            RichTextBox3.Enabled = True
            navigatedisable()
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Edit module: " & ex.Message)
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
            Label22.Text = "VIEW"
            navigateenable()
            LodaDataToTextBox()
        End If
        Exit Sub
    End Sub
    Private Sub insertData()
        Dim save As String
        If GrpAddCorrect = "C" Then
            save = "UPDATE [PARTY] SET P_CODE='" & TextBox2.Text & "',P_NAME='" & TextBox1.Text & "',AD1='" & RichTextBox1.Text.Replace("'", "''") & "',AD2=' ',AD3=' ',CITY='" & TextBox3.Text & "',HAD1='" & RichTextBox3.Text.Replace("'", "''") & "',HAD2=' ',HAD3=' ',HCT='" & TextBox5.Text & "',HPHON='" & TextBox6.Text & "',SPHON='" & TextBox4.Text & "',CONT_P='" & TextBox7.Text & "',STATE='" & TextBox8.Text & "',EMAIL_ID='" & TextBox9.Text & "',GST='" & TextBox10.Text & "' WHERE P_CODE='" & TextBox2.Text & "'" ' sorry about that
        Else
            save = "INSERT INTO [PARTY](P_CODE,P_NAME,AD1,AD2,AD3,CITY,HAD1,HAD2,HAD3,HCT,HPHON,SPHON,CONT_P,STATE,EMAIL_ID,GST) VALUES('" & TextBox2.Text & "','" & TextBox1.Text & "','" & RichTextBox1.Text.Replace("'", "''") & "','" & " " & "','" & " " & "','" & TextBox3.Text & "','" & RichTextBox3.Text.Replace("'", "''") & "','" & " " & "','" & " " & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox4.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "')"
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
            MsgBox("Exception: Data Insertion in PARTY table in database" & ex.Message)
        End Try
    End Sub
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Dim errorMsg As String = "Please enter tenant's name"
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
            Dim errorMsg As String = "Please Enter tenant code"
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
            da = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] where [P_CODE]='" & Trim(TextBox2.Text) & "'", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "PARTY")

            If ds.Tables(0).Rows.Count > 0 And GrpAddCorrect <> "C" Then
                errorMsg = "Duplicate tenant code not allowed..."
                e.Cancel = True
                TextBox2.Select(0, TextBox2.Text.Length)
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
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        DataGridView1.Enabled = False
        MyConn = New OleDbConnection(connString)
        If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
        End If
        da = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] where [P_CODE]='" & Trim(TextBox2.Text) & "'", MyConn)
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
                    objcmd.CommandText = "delete from [PARTY] where P_CODE='" & Trim(TextBox2.Text) & "'"
                    objcmd.ExecuteNonQuery()
                    MsgBox("Data deleted successfully from PARTY table in database", vbInformation)
                    objcmd.Dispose()
                    MyConn.Close()
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    DataGridView1.Enabled = True
                    da = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] order by P_CODE", MyConn)
                    ds = New DataSet
                    ds.Clear()
                    da.Fill(ds, "PARTY")
                    DataGridView1.DataSource = ds.Tables("PARTY")
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
    Private Sub cmdCancel_MouseEnter(sender As Object, e As EventArgs) Handles cmdCancel.MouseEnter
        bValidatepcode = False
        bValidatepname = False
        BVALIDATEEMAIL = False
    End Sub

    Private Sub cmdCancel_MouseLeave(sender As Object, e As EventArgs) Handles cmdCancel.MouseLeave
        bValidatepcode = True
        bValidatepname = True
        BVALIDATEEMAIL = True

    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
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
            indexorder = "P_CODE"
            GroupBox3.Text = "Search by tenant code"
            DataGridView1.Sort(DataGridView1.Columns(0), SortOrder.Ascending)
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "P_NAME"
            GroupBox3.Text = "Search by tenant name"
            DataGridView1.Sort(DataGridView1.Columns(1), SortOrder.Ascending)
        End If
        LodaDataToTextBox()
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseDoubleClick
        If e.ColumnIndex = 0 Then
            indexorder = "P_CODE"
            GroupBox3.Text = "Search by tenant code"
            DataGridView1.Sort(DataGridView1.Columns(0), SortOrder.Ascending)
        End If
        If e.ColumnIndex = 1 Then
            indexorder = "P_NAME"
            GroupBox3.Text = "Search by tenant name"
            DataGridView1.Sort(DataGridView1.Columns(1), SortOrder.Ascending)
        End If
        LodaDataToTextBox()
    End Sub
    Private Sub TextBox9_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Dim errorMsg As String = "Please enter valid email address"
        If BVALIDATEEMAIL = True And TextBox9.Text.Trim <> "" Then
            ' Cancel the event and select the text to be corrected by the user.
            If EmailAddressCheck(TextBox9.Text) = False Then
                e.Cancel = True
                TextBox9.Select(0, TextBox9.Text.Length)
                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(TextBox9, errorMsg)
            End If

        End If
    End Sub

    Private Sub TextBox9_Validated(sender As Object, e As EventArgs) Handles TextBox9.Validated
        ErrorProvider1.SetError(TextBox9, "")
    End Sub
    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        'Dim pattern As String = "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Private Sub FrmTenant_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub TxtSrch_TextChanged(sender As Object, e As EventArgs) Handles TxtSrch.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class
