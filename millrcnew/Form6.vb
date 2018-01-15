Imports System.Data.OleDb
Imports System.Globalization
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO

Public Class FrmInvoicePrn
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
    Dim indexorder As String = "[PARTY].P_NAME"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Private Sub FrmInvoicePrn_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        ' GroupBox5.Visible = False
        '  DataGridView2.Enabled = False
        For Each column As DataGridViewColumn In DataGridView2.Columns

            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
        'ComboBox4.Text = DateTime.Now.Year
        If DateTime.Now.Month = 1 Then
            ComboBox3.Text = DateAndTime.MonthName(12)
            ComboBox4.Text = DateTime.Now.Year - 1
        Else
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox4.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
        End If
        TextBox1.Focus()
        formloaded = True
    End Sub
    Private Sub KeyUpHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.F1 Then
        End If
    End Sub

    Private Sub KeyDownHandler(ByVal o As Object, ByVal e As KeyEventArgs)
        e.SuppressKeyPress = True

    End Sub
    Private Sub FrmInvoicePrn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "TextBox1" Or Me.ActiveControl.Name = "TextBox2") Then
            DataGridView2.Visible = True

            GroupBox5.Visible = True
            Me.Width = Me.Width + DataGridView2.Width - 15
            Me.Height = Me.Height + 145
            ctrlname = Me.ActiveControl.Name
            ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        End If
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
    Private Sub ShowData(mnth As String, yr As String)
        '  konek() 'open our connection
        Try

            MyConn = New OleDbConnection(connString)
            'If MyConn.State = ConnectionState.Closed Then
            MyConn.Open()
            ' End If
            'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].GROUP,[BILL].GODWN_NO", MyConn)
            da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & mnth & "' AND YEAR([BILL].BILL_DATE)='" & yr & "' order by [BILL].BILL_DATE,[BILL].INVOICE_NO", MyConn)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for selected month")
            ComboBox3.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Invoice number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Invoice number")
            TextBox2.Focus()
            Exit Sub
        End If
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Invoice number must be less than To invoice number")
            Exit Sub
        End If

        Dim myList As New List(Of String)()

        Dim FLNAME As String = GetValue(DataGridView2.Item(0, Convert.ToInt32(TextBox3.Text)).Value).Replace("/", "_")   ' 'TextBox1.Text.Substring(0, 4) & "_" & Month(CDate("1 " & TextBox1.Text.Substring(8, 3))) & "_" & TextBox1.Text.Substring(12, TextBox1.Text.Length - 12)
        Form7.RichTextBox1.LoadFile(Application.StartupPath & "\Invoices\dat\" & ComboBox4.Text & "\" & ComboBox3.Text & "\" & FLNAME & ".dat", RichTextBoxStreamType.PlainText)
        Form7.Show()
        DataGridView2.Rows(Convert.ToInt32(TextBox3.Text)).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(Convert.ToInt32(TextBox3.Text)).Cells(0)
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount <= 1 Then
            MsgBox("No data exist for selected month")
            ComboBox3.Focus()
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Invoice number")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Invoice number")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strbill As Int32 = Convert.ToInt32(TextBox3.Text)
        Dim edbill As Int32 = Convert.ToInt32(TextBox4.Text)

        If strbill > edbill Then
            MsgBox("From Invoice number must be less than To invoice number")
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
        For X As Integer = strbill To edbill
            myList.Add(GetValue(DataGridView2.Item(0, X).Value).Replace("/", "_").Replace(" ", "_"))
            ' FILE_NO = FILE_NO.Replace(" ", "_")
        Next
        Dim myArray As String() = myList.ToArray()
        For X As Integer = 0 To myArray.Length - 1
            ' MsgBox(myArray(X))
            Dim strPDFFile As String = Dir(Application.StartupPath & "\Invoices\pdf\" & ComboBox4.Text & "\" & ComboBox3.Text & "\" & myArray(X) & ".pdf")
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
            MyProcess.StartInfo.FileName = Application.StartupPath & "\Invoices\pdf\" & ComboBox4.Text & "\" & ComboBox3.Text & "\" & strPDFFile
            MyProcess.Start()
            MyProcess.WaitForExit(10000)
            ' MyProcess.CloseMainWindow()
            MyProcess.Close()
            Threading.Thread.Sleep(5000)
            ' strPDFFile = Dir()
            ' Loop
        Next
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month)
        End If
        If Trim(ComboBox4.Text) = "" Then
            ComboBox4.Text = DateTime.Now.Year
        End If
        ShowData(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month, ComboBox4.Text)
        If DataGridView2.RowCount > 1 Then
            TextBox1.Text = DataGridView2.Item(0, 0).Value
            TextBox2.Text = DataGridView2.Item(0, DataGridView2.RowCount - 2).Value
            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, 0).Value.Substring(12, 3)) - 1
            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, DataGridView2.RowCount - 2).Value.Substring(12, 3)) - 1
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub


    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        'Dim i As Integer = DataGridView2.CurrentRow.Index
        'CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(0, i).Value)
        'If (TextBox2.Text = "") Then
        '    TextBox2.Text = GetValue(DataGridView2.Item(0, i).Value)
        'End If
        ''  GroupBox5.Visible = False
        '' DataGridView2.Visible = False
        ''  Me.Width = Me.Width - DataGridView2.Width + 15
        '' Me.Height = Me.Height - 145
        'If ctrlname = "TextBox1" Then
        '    TextBox2.Focus()
        'Else
        '    If ctrlname = "TextBox2" Then
        '        Button1.Focus()
        '    Else
        '        TextBox1.Focus()
        '    End If
        'End If
    End Sub

    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        Console.WriteLine(TxtSrch.Text)
        'da = New OleDb.OleDbDataAdapter("SELECT [BILL].*,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR([BILL].BILL_DATE)='" & ComboBox4.Text & "' and " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [BILL].GODWN_NO", MyConn)
        da = New OleDb.OleDbDataAdapter("SELECT [BILL].INVOICE_NO,[BILL].GROUP,[BILL].GODWN_NO,[BILL].P_CODE,[BILL].BILL_DATE,[BILL].BILL_AMOUNT,[BILL].CGST_RATE,[BILL].CGST_AMT,[BILL].SGST_RATE,[BILL].SGST_AMT,[BILL].NET_AMOUNT,[BILL].HSN,SRNO,[BILL].REC_NO,[BILL].REC_DATE,[PARTY].P_NAME from [BILL] INNER JOIN [PARTY] on [BILL].P_CODE=[PARTY].P_CODE where MONTH([BILL].bill_date)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR([BILL].BILL_DATE)='" & ComboBox4.Text & "' and " & indexorder & " Like '%" & TxtSrch.Text & "%' ORDER BY [BILL].BILL_DATE+[BILL].INVOICE_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0).DefaultView
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        If Application.OpenForms().OfType(Of Form7).Any Then
            Form7.Close()
        End If
    End Sub

    Private Sub DataGridView2_Click(sender As Object, e As EventArgs) Handles DataGridView2.Click
        Dim i As Integer = DataGridView2.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView2.Item(0, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView2.Item(0, i).Value)
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

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        ctrlname = "TextBox1"
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        ctrlname = "TextBox2"
    End Sub

    Private Sub FrmInvoicePrn_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub ComboBox4_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox4.TextUpdate
        If ComboBox4.FindString(ComboBox4.Text) < 0 Then
            ComboBox4.Text = ComboBox4.Text.Remove(ComboBox4.Text.Length - 1)
            ComboBox4.SelectionStart = ComboBox4.Text.Length

        End If
    End Sub
    Private Sub ComboBox3_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox3.TextUpdate
        If ComboBox3.FindString(ComboBox3.Text) < 0 Then
            ComboBox3.Text = ComboBox3.Text.Remove(ComboBox3.Text.Length - 1)
            ComboBox3.SelectionStart = ComboBox3.Text.Length

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Trim.Equals("") Then
        Else
            For i As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Rows(i).Cells(0).Value IsNot Nothing Then
                    '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
                    If Convert.ToInt32(DataGridView2.Rows(i).Cells(0).Value) > (Convert.ToInt32(TextBox1.Text)) Then
                        DataGridView2.ClearSelection()
                        If i = 0 Then
                            DataGridView2.Rows(i).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                        Else
                            DataGridView2.Rows(i - 1).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
                        End If
                        'DataGridView2.Rows(i).Cells(0).Selected = True
                        'DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        ''DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        'TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1  'DataGridView2.Rows(i).Cells(0).Value   'DataGridView2.SelectedCells.Item(0).Value
                        'MessageBox.Show(irowindex)
                        Exit For
                    Else
                        DataGridView2.Rows(i).Cells(0).Selected = True
                        DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Trim.Equals("") Then
        Else
            For i As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Rows(i).Cells(0).Value IsNot Nothing Then
                    '  If DataGridView2.Rows(i).Cells(0).Value.ToString.ToUpper.Contains(TextBox1.Text.ToUpper) Then
                    If Convert.ToInt32(DataGridView2.Rows(i).Cells(0).Value) > (Convert.ToInt32(TextBox2.Text)) Then
                        DataGridView2.ClearSelection()
                        If i = 0 Then
                            DataGridView2.Rows(i).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                        Else
                            DataGridView2.Rows(i - 1).Cells(0).Selected = True
                            DataGridView2.CurrentCell = DataGridView2.Rows(i - 1).Cells(0)
                            'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                            TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i - 1).Value.Substring(12, 3)) - 1
                        End If
                        'DataGridView2.Rows(i).Cells(0).Selected = True
                        'DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        ''DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        'TextBox3.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1  'DataGridView2.Rows(i).Cells(0).Value   'DataGridView2.SelectedCells.Item(0).Value
                        'MessageBox.Show(irowindex)
                        Exit For
                    Else
                        DataGridView2.Rows(i).Cells(0).Selected = True
                        DataGridView2.CurrentCell = DataGridView2.Rows(i).Cells(0)
                        'DataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DimGray
                        TextBox4.Text = Convert.ToInt32(DataGridView2.Item(12, i).Value.Substring(12, 3)) - 1
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

End Class