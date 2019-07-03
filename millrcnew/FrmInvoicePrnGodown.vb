Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
''' <summary>
''' tables used - bill, party,group,godown
''' this is form to accept inputs from user to view/print invoices
''' Form20.vb is used to hold report view
''' 
''' </summary>
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
        ''''''set position of the form
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True

        For Each column As DataGridViewColumn In DataGridView2.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable     '''''don't allow user to change sorting order of invoice datagrid by clicking on column headers
        Next
        fillgroupcombo()      '''''fill godown group combo box using group table and initially dont show any group seleted
        groupfilled = True
        ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
        ComboBox1.Text = ""
        fillgodowncombo()        '''''fill godown number combo box with godown number using godown table and don't show any godown selected
        godownfilled = True
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""
        ShowData(ComboBox1.Text, TextBox1.Text)       ''''fill invoice data grid with invoices falling in selected month-year using bill table
    End Sub

    Private Sub FrmInvoicePrnGodown_Move(sender As Object, e As EventArgs) Handles Me.Move
        ''''keep position of the form fix on MDI form
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
    Private Sub ShowData(grp As String, frGdn As String)
        ''''fill invoice data grid with invoices for group+godwn_no falling in selected month-year using bill table
        Try
            MyConn = New OleDbConnection(connString)
            MyConn.Open()
            If (ComboBox2.Text.Equals("")) Then     '''''if godown number not selected select all invoices for selected group ,month and year from bill table
                da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            Else        '''''if godown number selected select all invoices for selected group ,godown  umber,month and year from bill table
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
        '''''''''Report view
        If DataGridView2.RowCount < 1 Then
            MsgBox("No data exist for this godown")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (ComboBox2.Text = "") Then
            MsgBox("Please enter Godown number")
            ComboBox2.Focus()
            Exit Sub
        End If
        Dim myList As New List(Of String)()
        Dim myyrList As New List(Of String)()
        Dim mymnList As New List(Of String)()
        For X As Integer = 0 To DataGridView2.RowCount - 1
            If DataGridView2.Item(0, X).Value = True Then
                Dim yr As String = Year(DataGridView2.Item(5, X).Value)         '''''add year to year list
                Dim mnthname As String = MonthName(Month(DataGridView2.Item(5, X).Value), False)    '''''add month name to month list
                myList.Add(GetValue(DataGridView2.Item(1, X).Value).Replace("/", "_").Replace(" ", "_"))   ''''add invoice number to invoice list
                myyrList.Add(yr)
                mymnList.Add(mnthname)
            End If
        Next
        myArray = myList.ToArray()
        myYrArray = myyrList.ToArray()
        myMnArray = mymnList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Invoice")
            ComboBox2.Focus()
            Exit Sub
        End If
        ''''''''''File name for the starting invoice which is already generate using invoice generate process
        Dim FLNAME As String = myArray(0).Replace("/", "_")   ' 'TextBox1.Text.Substring(0, 4) & "_" & Month(CDate("1 " & TextBox1.Text.Substring(8, 3))) & "_" & TextBox1.Text.Substring(12, TextBox1.Text.Length - 12)
        Form20.Label1.Text = 0
        ''''''bind .dat invoice file in  richtextbox in another form used for report view
        Form20.RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & myyrList(0) & "\" & mymnList(0) & "\" & FLNAME & ".dat", RichTextBoxStreamType.PlainText)
        Form20.Show()    ''''show invoice report

    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        ''''select / unselect invoices using check box
        If (DataGridView2.CurrentCell.ColumnIndex = 0) Then
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
        ''''''fill godown number combo using godown table
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
        ''''''''if focus is on godown number combo boxand user press F1 key open help form
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "ComboBox2") Then
            helpgrpcombo = ComboBox1
            helpgdncombo = ComboBox2
            GodownHelp.Show()
        End If
    End Sub
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''''report print
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
        Dim myList As New List(Of String)()          '''''array to hold selected invoice numbers for selected group,godwn_no
        Dim myyrList As New List(Of String)()
        Dim mymnList As New List(Of String)()
        For X As Integer = 0 To DataGridView2.RowCount - 1
            If DataGridView2.Item(0, X).Value = True Then
                Dim yr As String = Year(DataGridView2.Item(5, X).Value)
                Dim mnthname As String = MonthName(Month(DataGridView2.Item(5, X).Value), False)
                myList.Add(GetValue(DataGridView2.Item(1, X).Value).Replace("/", "_").Replace(" ", "_"))
                myyrList.Add(yr)
                mymnList.Add(mnthname)
            End If
        Next
        myArray = myList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Invoice")
            ComboBox2.Focus()
            Exit Sub
        End If

        '''''looping through invoice array
        For X As Integer = 0 To myArray.Length - 1
            ''''''pdf filename
            Dim strPDFFile As String = Dir(Application.StartupPath & "\Invoices\pdf\" & myyrList(X) & "\" & mymnList(X) & "\" & myArray(X) & ".pdf")
            Dim PrintPDFFile As New ProcessStartInfo
            Dim MyProcess As New Process
            '''''''''print pdf file
            MyProcess.StartInfo.UseShellExecute = True
            MyProcess.StartInfo.CreateNoWindow = True
            MyProcess.StartInfo.Verb = "print"
            MyProcess.StartInfo.FileName = Application.StartupPath & "\Invoices\pdf\" & myyrList(X) & "\" & mymnList(X) & "\" & strPDFFile
            MyProcess.Start()
            MyProcess.WaitForExit(10000)
            MyProcess.Close()
            Threading.Thread.Sleep(5000)
        Next
    End Sub
    Public Function fillgroupcombo()
        ''''''fill godown group combo using group table
        Try
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ''''''when user select group from group combo box, fill godown number combobox with godown number of that group using godown table and show it in datagrid
        If groupfilled Then
            fillgodowncombo()
            ShowData(ComboBox1.Text, "")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()      ''''close form
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If godownfilled Then

            ShowData(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        ''''''''when user select godown number form godown combo, fill datagrid with invoices generated for that group+godwn_no from bill table
        If godownfilled Then
            ShowData(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub
End Class