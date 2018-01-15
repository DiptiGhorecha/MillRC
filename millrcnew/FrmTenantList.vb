Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmTenantList
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
    Dim indexorder As String = "P_CODE"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim fnum As Integer

    Private Sub FrmTenantList_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        'fillpartycombo(ComboBox1)
        'fillpartycombo(ComboBox2)
        ShowData()
        If DataGridView1.RowCount >= 1 Then
        End If
        ' TextBox6.Text = "1"
        ' TextBox1.Enabled = False
        ' TextBox2.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
        ComboBox1.Text = ""
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""
        formloaded = True
        If DataGridView1.RowCount > 1 Then
            TextBox1.Text = DataGridView1.Item(0, 0).Value
            TextBox2.Text = DataGridView1.Item(0, DataGridView1.RowCount - 1).Value
        End If
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
    Public Function fillpartycombo(cmbo As ComboBox)
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            da = New OleDb.OleDbDataAdapter("SELECT * from [PARTY] Order by [PARTY].P_NAME", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "PARTY")
            cmbo.DataSource = ds.Tables("PARTY")
            cmbo.DisplayMember = "P_NAME"
            cmbo.ValueMember = "P_CODE"
            da.Dispose()
            ds.Dispose()
            MyConn.Close() ' close connection

        Catch ex As Exception
            MessageBox.Show("Party combo fill :" & ex.Message)
        End Try
    End Function
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        Dim i As Integer = DataGridView1.CurrentRow.Index
        CType(Me.Controls.Find(ctrlname, False)(0), TextBox).Text = GetValue(DataGridView1.Item(0, i).Value)
        If (TextBox2.Text = "") Then
            TextBox2.Text = GetValue(DataGridView1.Item(0, i).Value)
        End If
        '  GroupBox5.Visible = False
        '  DataGridView2.Visible = False
        ' ' Me.Width = Me.Width - DataGridView2.Width + 15
        '  Me.Height = Me.Height - 145
        If ctrlname = "TextBox1" Then
            TextBox2.Focus()
        Else
            If ctrlname = "TextBox2" Then
                Button1.Focus()
            Else
                TextBox1.Focus()
            End If
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

    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        ctrlname = "TextBox1"
    End Sub
    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        ctrlname = "TextBox2"
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex <> -1 Then
            TextBox1.Text = ComboBox1.SelectedValue.ToString
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <> -1 Then
            TextBox2.Text = ComboBox2.SelectedValue.ToString
        End If
    End Sub

    Private Sub FrmTenantList_Move(sender As Object, e As EventArgs) Handles Me.Move
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Tenant")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Tenant")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As String = TextBox1.Text
        Dim edrec As String = TextBox2.Text

        If edrec < strrec Then
            MsgBox("From Tenant must be less than To Tenant")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Tenantmasterlist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If

        chkrs1.Open("SELECT * from [PARTY] where P_CODE>='" & strrec & "' AND P_code<='" & edrec & "' order by P_NAME", xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1
        Dim ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GST, REMARK As String
        Do While chkrs1.EOF = False


            If first Then
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(12, "Tenant Code", "S") & " " & GetStringToPrint(55, "Tenant Name", "S") & " " & GetStringToPrint(25, "Office Phone No", "S") & " " & GetStringToPrint(25, "Resi. Phone Number", "S") & " " & GetStringToPrint(55, "Contact Person", "S") & " " & GetStringToPrint(30, "Email_ID", "S") & " " & GetStringToPrint(15, "GST", "S") & vbNewLine)
                Print(fnum, StrDup(230, "=") & vbNewLine)
                first = False
                xline = xline + 5
            End If


            If IsDBNull(chkrs1.Fields(13).Value) Then   'Contact Person
                CPERSON = ""
            Else
                CPERSON = chkrs1.Fields(13).Value
            End If

            If IsDBNull(chkrs1.Fields(18).Value) Then   'Email Address
                EMAIL = ""
            Else
                EMAIL = chkrs1.Fields(18).Value
            End If

            If IsDBNull(chkrs1.Fields(19).Value) Then   'GST
                GST = ""
            Else
                GST = chkrs1.Fields(19).Value
            End If

            If IsDBNull(chkrs1.Fields(12).Value) Then   'Remark
                REMARK = ""
            Else
                REMARK = chkrs1.Fields(12).Value
            End If

            If IsDBNull(chkrs1.Fields(10).Value) Then    'Address & Phone No
                ADDPHONE = ""
            Else
                ADDPHONE = chkrs1.Fields(10).Value
            End If

            If IsDBNull(chkrs1.Fields(2).Value) Then
                ADD1 = ""
            Else
                ADD1 = chkrs1.Fields(2).Value
            End If

            If IsDBNull(chkrs1.Fields(3).Value) Then
                ADD2 = ""
            Else
                ADD2 = chkrs1.Fields(3).Value
            End If
            If IsDBNull(chkrs1.Fields(4).Value) Then
                ADD3 = ""
            Else
                ADD3 = chkrs1.Fields(4).Value
            End If
            If IsDBNull(chkrs1.Fields(5).Value) Then
                ACT = ""
            Else
                ACT = chkrs1.Fields(5).Value
            End If
            Dim addArr() As String
            If (ADD1.IndexOf(vbLf) >= 0) Then
                addArr = ADD1.Split(vbLf)
                ADD1 = addArr(0)
                If addArr.Length > 1 Then
                    ADD2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                        ADD3 = addArr(2)
                    End If
                End If

                If IsDBNull(chkrs1.Fields(11).Value) Then    'House Address & Phone No
                HPHONE = ""
            Else
                HPHONE = chkrs1.Fields(11).Value
            End If
            If IsDBNull(chkrs1.Fields(6).Value) Then
                HADD1 = ""
            Else
                HADD1 = chkrs1.Fields(6).Value     '.Replace(vbCr, "").Replace(vbLf, "") ' chkrs1.Fields(6).Value

            End If
            If IsDBNull(chkrs1.Fields(7).Value) Then
                HADD2 = ""
            Else
                HADD2 = chkrs1.Fields(7).Value
            End If
            If IsDBNull(chkrs1.Fields(8).Value) Then
                HADD3 = ""
            Else
                HADD3 = chkrs1.Fields(8).Value
            End If
            If IsDBNull(chkrs1.Fields(9).Value) Then
                HCT = ""
            Else
                HCT = chkrs1.Fields(9).Value
            End If
            Dim addHrr() As String
            If (HADD1.IndexOf(vbLf) >= 0) Then
                addHrr = HADD1.Split(vbLf)
                HADD1 = addHrr(0)
                If addHrr.Length > 1 Then
                    HADD2 = addHrr(1)
                End If
                If addHrr.Length > 2 Then
                    HADD3 = addHrr(2)
                End If
            End If


            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(12, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & " " & GetStringToPrint(25, ADDPHONE, "S") & " " & GetStringToPrint(25, HPHONE, "S") & " " & GetStringToPrint(55, CPERSON, "S") & " " & GetStringToPrint(30, EMAIL, "S") & " " & GetStringToPrint(15, GST, "S") & vbNewLine)
            'If (ADD2.Equals("")) Then
            '    If (ADD3.Equals("")) Then
            '        Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ACT, "S"))
            '    Else
            '        Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD3, "S"))
            '    End If
            'Else
            '    Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD2, "S"))
            'End If

            'If (HADD2.Equals("")) Then
            '    If (HADD3.Equals("")) Then
            '        Print(fnum, GetStringToPrint(75, HCT, "S"))
            '    Else
            '        Print(fnum, GetStringToPrint(75, HADD3, "S"))
            '    End If
            'Else
            '    Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD2, "S"))
            'End If
            ' Print(fnum, " " & vbNewLine)

            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FrmTenantListView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Tenantmasterlist.dat", RichTextBoxStreamType.PlainText)
        FrmTenantListView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Tenantmasterlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 15
            pdfPage.Width = 842
            pdfPage.Height = 595
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 6, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1
                line = readFile.ReadLine()
                If counter >= 43 Then
                    counter = 0
                    pdfPage = pdf.AddPage()
                    graph = XGraphics.FromPdfPage(pdfPage)
                    font = New XFont("COURIER NEW", 6, XFontStyle.Regular)

                    pdfPage.TrimMargins.Left = 15

                    pdfPage.Width = 842
                    pdfPage.Height = 595
                    yPoint = 60
                End If
                If line Is Nothing Then
                    Exit While
                Else
                    graph.DrawString(line, font, XBrushes.Black,
                    New XRect(30, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = invoice_no & ".pdf"


            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (TextBox1.Text = "") Then
            MsgBox("Please enter From Tenant")
            TextBox1.Focus()
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Please enter To Tenant")
            TextBox2.Focus()
            Exit Sub
        End If

        Dim strrec As String = TextBox1.Text
        Dim edrec As String = TextBox2.Text

        If edrec < strrec Then
            MsgBox("From Tenant must be less than To Tenant")
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\Tenantmasterlist.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If

        chkrs1.Open("SELECT * from [PARTY] where P_CODE>='" & strrec & "' AND P_code<='" & edrec & "' order by P_NAME", xcon)
        If chkrs1.BOF = False Then
            chkrs1.MoveFirst()
        End If
        Dim first As Boolean = True
        Dim counter As Integer = 1
        Dim ADDPHONE, ADD1, ADD2, ADD3, ACT, HPHONE, HADD1, HADD2, HADD3, HCT, CPERSON, EMAIL, GST, REMARK As String
        Do While chkrs1.EOF = False


            If first Then
                Print(fnum, GetStringToPrint(7, "Sr. No.", "S") & " " & GetStringToPrint(12, "Tenant Code", "S") & " " & GetStringToPrint(55, "Tenant Name", "S") & " " & GetStringToPrint(25, "Office Phone No", "S") & " " & GetStringToPrint(25, "Resi. Phone Number", "S") & " " & GetStringToPrint(55, "Contact Person", "S") & " " & GetStringToPrint(30, "Email_ID", "S") & " " & GetStringToPrint(15, "GST", "S") & vbNewLine)
                Print(fnum, StrDup(231, "=") & vbNewLine)
                first = False
                xline = 2
            End If


            If IsDBNull(chkrs1.Fields(13).Value) Then   'Contact Person
                CPERSON = ""
            Else
                CPERSON = chkrs1.Fields(13).Value
            End If

            If IsDBNull(chkrs1.Fields(18).Value) Then   'Email Address
                EMAIL = ""
            Else
                EMAIL = chkrs1.Fields(18).Value
            End If

            If IsDBNull(chkrs1.Fields(19).Value) Then   'GST
                GST = ""
            Else
                GST = chkrs1.Fields(19).Value
            End If

            If IsDBNull(chkrs1.Fields(12).Value) Then   'Remark
                REMARK = ""
            Else
                REMARK = chkrs1.Fields(12).Value
            End If

            If IsDBNull(chkrs1.Fields(10).Value) Then    'Address & Phone No
                ADDPHONE = ""
            Else
                ADDPHONE = chkrs1.Fields(10).Value
            End If

            If IsDBNull(chkrs1.Fields(2).Value) Then
                ADD1 = ""
            Else
                ADD1 = chkrs1.Fields(2).Value
            End If

            If IsDBNull(chkrs1.Fields(3).Value) Then
                ADD2 = ""
            Else
                ADD2 = chkrs1.Fields(3).Value
            End If
            If IsDBNull(chkrs1.Fields(4).Value) Then
                ADD3 = ""
            Else
                ADD3 = chkrs1.Fields(4).Value
            End If
            If IsDBNull(chkrs1.Fields(5).Value) Then
                ACT = ""
            Else
                ACT = chkrs1.Fields(5).Value
            End If
            Dim addArr() As String
            If (ADD1.IndexOf(vbLf) >= 0) Then
                addArr = ADD1.Split(vbLf)
                ADD1 = addArr(0)
                If addArr.Length > 1 Then
                    ADD2 = addArr(1)
                End If
                If addArr.Length > 2 Then
                    ADD3 = addArr(2)
                End If
            End If

            If IsDBNull(chkrs1.Fields(11).Value) Then    'House Address & Phone No
                HPHONE = ""
            Else
                HPHONE = chkrs1.Fields(11).Value
            End If
            If IsDBNull(chkrs1.Fields(6).Value) Then
                HADD1 = ""
            Else
                HADD1 = chkrs1.Fields(6).Value     '.Replace(vbCr, "").Replace(vbLf, "") ' chkrs1.Fields(6).Value

            End If
            If IsDBNull(chkrs1.Fields(7).Value) Then
                HADD2 = ""
            Else
                HADD2 = chkrs1.Fields(7).Value
            End If
            If IsDBNull(chkrs1.Fields(8).Value) Then
                HADD3 = ""
            Else
                HADD3 = chkrs1.Fields(8).Value
            End If
            If IsDBNull(chkrs1.Fields(9).Value) Then
                HCT = ""
            Else
                HCT = chkrs1.Fields(9).Value
            End If
            Dim addHrr() As String
            If (HADD1.IndexOf(vbLf) >= 0) Then
                addHrr = HADD1.Split(vbLf)
                HADD1 = addHrr(0)
                If addHrr.Length > 1 Then
                    HADD2 = addHrr(1)
                End If
                If addHrr.Length > 2 Then
                    HADD3 = addHrr(2)
                End If
            End If


            Print(fnum, GetStringToPrint(7, counter, "S") & " " & GetStringToPrint(12, chkrs1.Fields(0).Value, "S") & " " & GetStringToPrint(55, chkrs1.Fields(1).Value, "S") & " " & GetStringToPrint(25, ADDPHONE, "S") & " " & GetStringToPrint(25, HPHONE, "S") & " " & GetStringToPrint(55, CPERSON, "S") & " " & GetStringToPrint(30, EMAIL, "S") & " " & GetStringToPrint(15, GST, "S") & vbNewLine)
            'If (ADD2.Equals("")) Then
            '    If (ADD3.Equals("")) Then
            '        Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ACT, "S"))
            '    Else
            '        Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD3, "S"))
            '    End If
            'Else
            '    Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD2, "S"))
            'End If

            'If (HADD2.Equals("")) Then
            '    If (HADD3.Equals("")) Then
            '        Print(fnum, GetStringToPrint(75, HCT, "S"))
            '    Else
            '        Print(fnum, GetStringToPrint(75, HADD3, "S"))
            '    End If
            'Else
            '    Print(fnum, GetStringToPrint(10, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(55, "", "S") & GetStringToPrint(75, ADD2, "S"))
            'End If
            ' Print(fnum, " " & vbNewLine)
            xline = xline + 1
            counter = counter + 1
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If

        Loop
        chkrs1.Close()
        MyConn.Close()
        FileClose(fnum)
        FrmTenantListView.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Tenantmasterlist.dat", RichTextBoxStreamType.PlainText)
        FrmTenantListView.Show()
        CreatePDF(Application.StartupPath & "\Reports\Tenantmasterlist.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)

        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub
End Class