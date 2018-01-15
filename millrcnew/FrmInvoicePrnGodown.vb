Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmInvoicePrnGodown
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
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim dagp As OleDbDataAdapter
    Dim dsgp As DataSet
    Dim indexorder As String = "[PARTY].P_NAME"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim groupfilled As Boolean = False
    Dim godownfilled As Boolean = False

    Private Sub FrmInvoicePrnGodown_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        ' GroupBox5.Visible = False
        '  DataGridView2.Enabled = False
        For Each column As DataGridViewColumn In DataGridView2.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        fillgroupcombo()
        groupfilled = True
        ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
        ComboBox1.Text = ""
        fillgodowncombo()
        godownfilled = True
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""
        ShowData(ComboBox1.Text, TextBox1.Text)
        ' DataGridView2.EditMode = DataGridViewEditMode.EditOnEnter
    End Sub

    Private Sub FrmInvoicePrnGodown_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
    Private Sub ShowData(grp As String, frGdn As String)
        '  konek() 'open our connection
        Try

            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
            If (ComboBox2.Text.Equals("")) Then
                da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            Else
                da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' AND [BILL].GODWN_NO='" & ComboBox2.Text & "' order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            End If


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
            Dim chk As New DataGridViewCheckBoxColumn()
            chk.HeaderText = "Select Bill"
            chk.Name = "chk"
            chk.ValueType = GetType(Boolean)
            chk.DataPropertyName = "checker"
            DataGridView2.Columns.Insert(0, chk)
            DataGridView2.Columns(0).Width = 65
            DataGridView2.Columns(0).ReadOnly = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for this godown")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (ComboBox2.Text = "") Then
            MsgBox("Please enter Godown number")
            ComboBox2.Focus()
            Exit Sub
        End If
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim myList As New List(Of String)()
        Dim myyrList As New List(Of String)()
        Dim mymnList As New List(Of String)()
        For X As Integer = 0 To DataGridView2.RowCount - 1
            If DataGridView2.Item(0, X).Value = True Then
                Dim yr As String = Year(DataGridView2.Item(5, X).Value)
                Dim mnthname As String = MonthName(Month(DataGridView2.Item(5, X).Value), False)
                myList.Add(GetValue(DataGridView2.Item(1, X).Value).Replace("/", "_").Replace(" ", "_"))
                myyrList.Add(yr)
                mymnList.Add(mnthname)
                ' FILE_NO = FILE_NO.Replace(" ", "_")
            End If
        Next
        myArray = myList.ToArray()
        myYrArray = myyrList.ToArray()
        myMnArray = mymnList.ToArray()
        For X As Integer = 0 To myArray.Length - 1

        Next
        'Dim FLNAME As String = myyrList(0) & "\" & mymnList(0) & "\" & myArray(0).Replace("/", "_")   
        Dim FLNAME As String = myArray(0).Replace("/", "_")   ' 'TextBox1.Text.Substring(0, 4) & "_" & Month(CDate("1 " & TextBox1.Text.Substring(8, 3))) & "_" & TextBox1.Text.Substring(12, TextBox1.Text.Length - 12)
        Form20.Label1.Text = 0
        Form20.RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & myyrList(0) & "\" & mymnList(0) & "\" & FLNAME & ".dat", RichTextBoxStreamType.PlainText)
        Form20.Show()

    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        '  DataGridView2(e.ColumnIndex, e.RowIndex).[ReadOnly] = True
        If (DataGridView2.CurrentCell.ColumnIndex = 0) Then
            '  DataGridView2(e.ColumnIndex, e.RowIndex).[ReadOnly] = False
            'DataGridView2.BeginEdit(False)

            Dim CheckState As Boolean = DataGridView2.CurrentCell.Value
            If (CheckState = False) Then
                DataGridView2.CurrentCell.Value = True
            Else
                DataGridView2.CurrentCell.Value = False
            End If

        End If
    End Sub
    Private Sub KeyUpHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.F1 Then
        End If
    End Sub

    Private Sub KeyDownHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True

    End Sub
    Public Function fillgodowncombo()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            dagp = New OleDb.OleDbDataAdapter("SELECT * from [GODOWN] WHERE [GROUP]='" & ComboBox1.Text & "' and [STATUS]='C' Order by GODWN_NO", MyConn)
            dsgp = New DataSet
            dsgp.Clear()
            dagp.Fill(dsgp, "GODOWN")
            ComboBox2.DataSource = dsgp.Tables("GODOWN")
            ComboBox2.DisplayMember = "GODWN_NO"
            ComboBox2.ValueMember = "GODWN_NO"
            dagp.Dispose()
            dsgp.Dispose()
            MyConn.Close() ' close connection
        Catch ex As Exception
            MessageBox.Show("Godown combo fill :" & ex.Message)
        End Try
    End Function
    Private Sub FrmInvoicePrnGodown_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "ComboBox2") Then
            'GroupBox2.BringToFront()
            'GroupBox2.Visible = True
            'Label17.Text = "Godown Detail"
            'ctrlname = Me.ActiveControl.Name
            'GroupBox2.Focus()
            helpgrpcombo = ComboBox1
            helpgdncombo = ComboBox2
            GodownHelp.Show()

            '  ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        End If
    End Sub
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for this godown")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (ComboBox2.Text = "") Then
            MsgBox("Please enter Godown number")
            ComboBox2.Focus()
            Exit Sub
        End If
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim myList As New List(Of String)()
        Dim myyrList As New List(Of String)()
        Dim mymnList As New List(Of String)()
        For X As Integer = 0 To DataGridView2.RowCount - 1
            If DataGridView2.Item(0, X).Value = True Then
                Dim yr As String = Year(DataGridView2.Item(5, X).Value)
                Dim mnthname As String = MonthName(Month(DataGridView2.Item(5, X).Value), False)
                myList.Add(GetValue(DataGridView2.Item(1, X).Value).Replace("/", "_").Replace(" ", "_"))
                myyrList.Add(yr)
                mymnList.Add(mnthname)
                ' FILE_NO = FILE_NO.Replace(" ", "_")
            End If
        Next
        myArray = myList.ToArray()
        For X As Integer = 0 To myArray.Length - 1
            ' MsgBox(myArray(X))
            Dim strPDFFile As String = Dir(Application.StartupPath & "\Invoices\pdf\" & myyrList(X) & "\" & mymnList(X) & "\" & myArray(X) & ".pdf")
            Dim PrintPDFFile As New ProcessStartInfo
            Dim MyProcess As New Process
            MyProcess.StartInfo.UseShellExecute = True
            MyProcess.StartInfo.CreateNoWindow = True
            '  MyProcess.StartInfo.CreateNoWindow = False
            MyProcess.StartInfo.Verb = "print"
            MyProcess.StartInfo.FileName = Application.StartupPath & "\Invoices\pdf\" & myyrList(X) & "\" & mymnList(X) & "\" & strPDFFile
            MyProcess.Start()
            MyProcess.WaitForExit(10000)
            ' MyProcess.CloseMainWindow()
            MyProcess.Close()
            Threading.Thread.Sleep(5000)
        Next
    End Sub
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
        Catch ex As Exception
            MessageBox.Show("Group combo fill :" & ex.Message)
        End Try
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' ShowData(ComboBox1.Text, "")
        If groupfilled Then
            fillgodowncombo()
            ShowData(ComboBox1.Text, "")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ShowData(ComboBox1.Text, TextBox1.Text)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If godownfilled Then

            ShowData(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If godownfilled Then

            ShowData(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub
End Class