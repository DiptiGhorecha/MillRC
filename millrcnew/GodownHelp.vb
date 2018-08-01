Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Globalization
Public Class GodownHelp
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim rownum As Integer = 0
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim formloaded As Boolean
    Dim indexorder As String = "[PARTY].P_NAME"
    Function fillgroupbox2()
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [STATUS]='C' AND [GROUP]='" & helpgrpcombo.Text & "' order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [STATUS]='C' order by [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds)
            DataGridView3.DataSource = ds.Tables(0).DefaultView
            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            DataGridView3.Columns(1).Visible = False
            DataGridView3.Columns(2).Visible = False
            DataGridView3.Columns(4).Visible = False
            DataGridView3.Columns(5).Visible = False
            DataGridView3.Columns(6).Visible = False
            DataGridView3.Columns(7).Visible = False
            DataGridView3.Columns(8).Visible = False
            DataGridView3.Columns(9).Visible = False
            DataGridView3.Columns(10).Visible = False
            DataGridView3.Columns(11).Visible = False
            DataGridView3.Columns(12).Visible = False
            DataGridView3.Columns(13).Visible = False
            DataGridView3.Columns(14).Visible = False
            DataGridView3.Columns(15).Visible = False
            DataGridView3.Columns(16).Visible = False
            DataGridView3.Columns(17).Visible = False
            DataGridView3.Columns(18).Visible = False
            DataGridView3.Columns(19).Visible = False
            DataGridView3.Columns(20).Visible = False
            DataGridView3.Columns(21).Visible = False
            DataGridView3.Columns(22).Visible = False
            DataGridView3.Columns(23).Visible = False
            DataGridView3.Columns(24).Visible = False
            DataGridView3.Columns(25).Visible = False
            DataGridView3.Columns(26).Visible = False
            DataGridView3.Columns(27).Visible = False
            DataGridView3.Columns(28).Visible = False
            DataGridView3.Columns(29).Visible = False
            DataGridView3.Columns(30).Visible = False
            DataGridView3.Columns(31).Visible = False
            DataGridView3.Columns(32).Visible = False
            DataGridView3.Columns(33).Visible = False
            DataGridView3.Columns(34).Visible = False
            DataGridView3.Columns(35).Visible = False
            DataGridView3.Columns(36).Visible = False
            DataGridView3.Columns(37).Visible = False
            DataGridView3.Columns(0).Visible = True
            DataGridView3.Columns(3).Visible = True
            DataGridView3.Columns(38).Visible = True
            DataGridView3.Columns(0).HeaderText = "Group"
            DataGridView3.Columns(0).Width = 51
            DataGridView3.Columns(3).Width = 71
            DataGridView3.Columns(38).Width = 405
            DataGridView3.Columns(3).HeaderText = "Godown"
            DataGridView3.Columns(38).HeaderText = "Tenant"
            DataGridView3.Columns(21).HeaderText = "Outstanding"
            DataGridView3.Columns(21).Width = 105
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub GodownHelp_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = FormReceipt.Right + 10
        Me.KeyPreview = True
        fillgroupbox2()
        formloaded = True
        GodownHelp_Move(Nothing, Nothing)
        TxtSrch.Focus()
    End Sub
    Private Sub DataGridView3_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView3.DoubleClick
        helpgrpcombo.Text = FormReceipt.GetValue(DataGridView3.Item(0, DataGridView3.CurrentRow.Index).Value)
        helpgrpcombo.SelectedIndex = helpgrpcombo.FindStringExact(helpgrpcombo.Text)
        helpgdncombo.Text = FormReceipt.GetValue(DataGridView3.Item(3, DataGridView3.CurrentRow.Index).Value)
        helpgdncombo.SelectedIndex = helpgdncombo.FindStringExact(helpgdncombo.Text)

        Me.Close()
        'FormReceipt.Label14.Text = DataGridView3.Item(38, DataGridView3.CurrentRow.Index).Value
        'If DataGridView3.Item(37, DataGridView3.CurrentRow.Index).Value.Equals("997212") Then
        '    FormReceipt.Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
        'Else
        '    FormReceipt.Label15.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
        'End If

        '  GroupBox1.Visible = False
    End Sub
    Private Sub DataGridView3_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then
            indexorder = "[GODOWN].GROUP"
            GroupBox5.Text = "Search by Group Type"
            '    DataGridView2.Sort(DataGridView2.Columns(0), SortOrder.Descending)
        End If
        If e.ColumnIndex = 3 Then
            indexorder = "[GODOWN].GODWN_NO"
            GroupBox5.Text = "Search by Godown"
            '   DataGridView2.Sort(DataGridView2.Columns(3), SortOrder.Descending)
        End If
        If e.ColumnIndex = 38 Then
            indexorder = "[PARTY].P_NAME"
            GroupBox5.Text = "Search by tenant name"
            ' DataGridView2.Sort(DataGridView2.Columns(38), SortOrder.Descending)
        End If
    End Sub
    Private Sub TxtSrch_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtSrch.KeyUp
        MyConn = New OleDbConnection(connString)
        'If MyConn.State = ConnectionState.Closed Then
        MyConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where " & indexorder & " Like '%" & TxtSrch.Text & "%' and [STATUS]='C' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        'da = New OleDb.OleDbDataAdapter("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE where [GODOWN].GROUP Like '%" & TxtSrch.Text & "%' ORDER BY [GODOWN].GROUP+[GODOWN].GODWN_NO", MyConn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "GODOWN")
        DataGridView3.DataSource = ds.Tables("GODOWN")
        da.Dispose()
        ds.Dispose()
        MyConn.Close() ' clouse connection
        If e.KeyCode = 13 Then
            helpgdncombo.Text = FormReceipt.GetValue(DataGridView3.Item(3, DataGridView3.CurrentRow.Index).Value)
            helpgdncombo.SelectedIndex = helpgdncombo.FindStringExact(helpgdncombo.Text)
            Me.Close()
        End If
    End Sub
    Private Sub GodownHelp_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

End Class