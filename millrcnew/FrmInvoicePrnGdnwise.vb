Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmInvoicePrnGdnwise
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
    Dim indexorder As String = "[PARTY].P_NAME"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False

    Private Sub FrmInvoicePrnGdnwise_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        ShowData(ComboBox1.Text, TextBox1.Text, TextBox2.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(2, 0).Value
            TextBox2.Text = DataGridView2.Item(2, DataGridView2.RowCount - 2).Value
            TextBox3.Text = 0
            TextBox4.Text = DataGridView2.RowCount - 2
        End If
        TextBox1.Focus()
        formloaded = True
    End Sub
    Private Sub ShowData(grp As String, frGdn As String, toGdn As String)
        '  konek() 'open our connection
        Try

            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
            '  If (frGdn.Equals("") And toGdn.Equals("")) Then
            da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            '  Else
            '   da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' AND [BILL].GODWN_NO>='" & TextBox3.Text & "' AND [BILL].GODWN_NO<='" & TextBox4.Text & "' order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
            '   End If


            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0).DefaultView

            da.Dispose()
            ds.Dispose()
            MyConn.Close()
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'If TextBox1.Text.Trim.Equals("") Then
        'Else
        '    For i As Integer = 0 To DataGridView2.RowCount - 1
        '        If DataGridView2.Rows(i).Cells(2).Value IsNot Nothing Then
        '            '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
        '            If Convert.ToInt32(DataGridView2.Rows(i).Cells(2).Value) > (Convert.ToInt32(TextBox1.Text)) Then
        '                DataGridView2.ClearSelection()
        '                If i = 0 Then
        '                    DataGridView2.Rows(i).Cells(2).Selected = True
        '                    DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(2)
        '                    'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                    TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
        '                Else
        '                    DataGridView2.Rows(i - 1).Cells(2).Selected = True
        '                    DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(2)
        '                    'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                    TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
        '                End If
        '                Exit For
        '            Else
        '                DataGridView2.Rows(i).Cells(2).Selected = True
        '                DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(2)
        '                'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
        '            End If
        '        End If
        '    Next

        'End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'If TextBox2.Text.Trim.Equals("") Then
        'Else
        '    For i As Integer = 0 To DataGridView2.RowCount - 1
        '        If DataGridView2.Rows(i).Cells(2).Value IsNot Nothing Then
        '            '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
        '            If Convert.ToInt32(DataGridView2.Rows(i).Cells(2).Value) > (Convert.ToInt32(TextBox2.Text)) Then
        '                DataGridView2.ClearSelection()
        '                If i = 0 Then
        '                    DataGridView2.Rows(i).Cells(2).Selected = True
        '                    DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(2)
        '                    'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                    TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
        '                Else
        '                    DataGridView2.Rows(i - 1).Cells(2).Selected = True
        '                    DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(2)
        '                    'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                    TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
        '                End If
        '                Exit For
        '            Else
        '                DataGridView2.Rows(i).Cells(2).Selected = True
        '                DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(2)
        '                'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
        '                TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
        '            End If
        '        End If
        '    Next

        'End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        Dim i As Integer = DataGridView2.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(2, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView2.Item(2, i).Value)
            TextBox4.Text = DataGridView2.CurrentCell.RowIndex
        End If
        '  GroupBox5.Visible = False
        '  DataGridView2.Visible = False
        ' ' Me.Width = Me.Width - DataGridView2.Width + 15
        '  Me.Height = Me.Height - 145
        If ctrlname = "TextBox1" Then
            TextBox3.Text = DataGridView2.CurrentCell.RowIndex
            TextBox2.Focus()
        Else
            If ctrlname = "TextBox2" Then
                TextBox4.Text = DataGridView2.CurrentCell.RowIndex
                Button1.Focus()
            Else
                TextBox1.Focus()
            End If
        End If
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        ctrlname = "TextBox1"
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        ctrlname = "TextBox2"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ShowData(ComboBox1.Text, "", "")
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for selected month")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Godown number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Godown number")
            TextBox2.Focus()
            Exit Sub
        End If
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Godown number must be less than To Godown number")
            Exit Sub
        End If

        Dim myList As New List(Of String)()

        Dim FLNAME As String = GetValue(DataGridView2.Item(0, Convert.ToInt32(TextBox3.Text)).Value).Replace("/", "_")   ' 'TextBox1.Text.Substring(0, 4) & "_" & Month(CDate("1 " & TextBox1.Text.Substring(8, 3))) & "_" & TextBox1.Text.Substring(12, TextBox1.Text.Length - 12)
        Form19.RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & FLNAME & ".dat", RichTextBoxStreamType.PlainText)
        Form19.Show()
        DataGridView2.Rows(Convert.ToInt32(TextBox3.Text)).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(Convert.ToInt32(TextBox3.Text)).Cells(2)
    End Sub
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for selected month")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Godown number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Godown number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Godown number must be less than To Godown number")
            Exit Sub
        End If
        ' MsgBox(startBill1.Substring(0, 7))
        Me.PrintDialog1.PrintToFile = False
        If Me.PrintDialog1.ShowDialog() = DialogResult.OK Then
            '  Form7.PrintDocument1.PrinterSettings = Form7.PrintDialog1.PrinterSettings
            '  Form7.PrintDocument1.Print()
        End If
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        Dim myList As New List(Of String)()
        Dim myyrList As New List(Of String)()
        Dim mymnList As New List(Of String)()
        For X As Integer = strbill To edbill
            Dim yr As String = Year(DataGridView2.Item(4, X).Value)
            Dim mnthname As String = MonthName(Month(DataGridView2.Item(4, X).Value), False)
            myList.Add(GetValue(DataGridView2.Item(0, X).Value).Replace("/", "_").Replace(" ", "_"))
            myyrList.Add(yr)
            mymnList.Add(mnthname)
            ' FILE_NO = FILE_NO.Replace(" ", "_")
        Next
        Dim myArray As String() = myList.ToArray()
        For X As Integer = 0 To myArray.Length - 1
            ' MsgBox(myArray(X))
            Dim strPDFFile As String = Dir(Application.StartupPath & "\Invoices\pdf\" & myyrList(X) & "\" & mymnList(X) & "\" & myArray(X) & ".pdf")
            Dim PrintPDFFile As New ProcessStartInfo

            '  Do Until strPDFFile Is Nothing
            'PrintPDFFile.UseShellExecute = True
            'PrintPDFFile.Verb = "print"
            'PrintPDFFile.WindowStyle = ProcessWindowStyle.Hidden
            'PrintPDFFile.FileName = Application.StartupPath & "\Invoices\pdf\" & ComboBox4.Text & "\" & ComboBox3.Text & "\" & strPDFFile
            'Process.Start(PrintPDFFile)
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
            ' strPDFFile = Dir()
            ' Loop
        Next
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        Console.WriteLine(TxtSrch.Text)
        '  Dim daa As String = "Select Case [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] On [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' and " & indexorder & " Like '%" & TxtSrch.Text & "%'  order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO"
        da = New OleDb.OleDbDataAdapter("Select [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] On [BILL].P_CODE=[PARTY].P_CODE where [BILL].[GROUP]='" & ComboBox1.Text & "' and " & indexorder & " Like '%" & TxtSrch.Text & "%'  order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)

        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
    End Sub
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        TxtSrch.Text = ""

        If e.ColumnIndex = 1 Then
            indexorder = "[BILL].GROUP"
            GroupBox5.Text = "Search by Group"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 2 Then
            indexorder = "[BILL].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 15 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

End Class