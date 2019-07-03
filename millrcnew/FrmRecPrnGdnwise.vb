Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
''' <summary>
''' tables used - receipt, party, group ,godown, bill, gst
''' this is form to accept inputs from user to view/print receipts
''' Form21.vb is used to hold report view
''' </summary>
Public Class FrmRecPrnGdnwise
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim chkrs6 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dagp As OleDbDataAdapter
    Dim dsgp As DataSet
    Dim dag As OleDbDataAdapter
    Dim dsg As DataSet
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim fnum As Integer
    Dim groupfilled As Boolean = False
    Dim godownfilled As Boolean = False

    Private Sub FrmRecPrnGdnwise_Load(sender As Object, e As EventArgs) Handles Me.Load
        ''''''set position of the form
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        fillgroupcombo()    ''''''fill godown group combobox using group table and clear selection
        groupfilled = True
        ComboBox1.SelectedIndex = ComboBox1.Items.IndexOf("")
        ComboBox1.Text = ""
        fillgodowncombo()      ''''''fill godown number combobox using godown table and clear selection
        godownfilled = True
        ComboBox2.SelectedIndex = ComboBox2.Items.IndexOf("")
        ComboBox2.Text = ""
        showdata(ComboBox1.Text, ComboBox2.Text)  ''''fill receipt data grid with receipt for selected criteria using receipt table
    End Sub
    Public Function showdata(grp As String, gdn As String)
        ''''fill receipt data grid with receipt for selected criteria using receipt table
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            If (ComboBox2.Text.Equals("")) Then
                da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].[GROUP]='" & grp & "' order by [RECEIPT].REC_date,[RECEIPT].REC_NO", MyConn)
            Else
                da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].[GROUP]='" & grp & "' AND [RECEIPT].GODWN_NO='" & gdn & "' order by [RECEIPT].REC_date,[RECEIPT].REC_NO", MyConn)
            End If
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, "RECEIPT")
            DataGridView1.DataSource = ds.Tables("RECEIPT")
            da.Dispose()
            ds.Dispose()
            MyConn.Close()
            If DataGridView1.Columns.Contains("chk") Then
                DataGridView1.Columns.Remove("chk")
            End If
            DataGridView1.Columns(0).Visible = False
            Dim chk As New DataGridViewCheckBoxColumn()
            chk.HeaderText = "Select Receipt"
            chk.Name = "chk"
            chk.ValueType = GetType(Boolean)
            chk.DataPropertyName = "checker"
            DataGridView1.Columns.Insert(0, chk)
            DataGridView1.Columns(0).Width = 70
            DataGridView1.Columns(0).ReadOnly = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Public Function fillgodowncombo()
        ''''''fill godown combo with godown number using godown table
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
        ''''when user select group, fill godown number combo for that group and show receipt data to datagrid
        If groupfilled Then
            fillgodowncombo()
            showdata(ComboBox1.Text, "")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()  '''''close form
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ''''when user select godown number from godown number combo show receipt data to datagrid
        If godownfilled Then
            showdata(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        '''''''check/uncheck check boxes in datagrid
        If (DataGridView1.CurrentCell.ColumnIndex = 0) Then
            Dim CheckState As Boolean = DataGridView1.CurrentCell.Value
            If (CheckState = False) Then
                DataGridView1.CurrentCell.Value = True
            Else
                DataGridView1.CurrentCell.Value = False
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''''report view
        If DataGridView1.RowCount <= 0 Then
            MsgBox("No data exist for this godown")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (ComboBox2.Text = "") Then
            MsgBox("Please enter Godown number")
            ComboBox2.Focus()
            Exit Sub
        End If
        Dim recNoList As New List(Of String)()
        Dim recDateList As New List(Of Date)()
        Dim xline As Integer = 0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, X).Value = True Then
                recNoList.Add(GetValue(DataGridView1.Item(5, X).Value))
                recDateList.Add(Convert.ToDateTime(DataGridView1.Item(4, X).Value))
            End If
        Next
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports\")
        End If

        myArray = recNoList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Receipt")
            ComboBox2.Focus()
            Exit Sub
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprint1.dat", OpenMode.Output)
        Dim numRec As Integer = 0
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        For X As Integer = 0 To myArray.Length - 1
            Dim strn As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE=Format('" & Convert.ToDateTime(recDateList(X)) & "','dd/mm/yyyy') AND REC_NO=" & recNoList(X) & " order by [RECEIPT].REC_NO"
            chkrs1.Open(strn, xcon)
            If chkrs1.EOF = False Then
                chkrs1.MoveFirst()
            End If
            Dim advance As Double = chkrs1.Fields(5).Value
            For xX As Integer = 1 To Convert.ToInt16(TextBox6.Text.Trim)
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(17, "Receipt No.: GST-", "S") & GetStringToPrint(5, Trim(recNoList(X)), "S") & GetStringToPrint(41, " ", "S") & GetStringToPrint(7, "Date : ", "S") & GetStringToPrint(10, Trim(recDateList(X)), "S") & vbNewLine)
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = 6
                ''''''''''''''''''''''godown detail start
                chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & ComboBox1.Text & "' AND GODWN_NO='" & ComboBox2.Text & "' AND [STATUS]='C' ", xcon)
                Dim census As String = ""
                Dim survey As String = ""
                Dim pname As String = ""
                Dim rent As Double = 0
                Dim pcode1 As String = ""
                Dim hsnm As String = ""
                Dim CGST_RATE As Double = 0
                Dim SGST_RATE As Double = 0
                Dim CGST_TAXAMT As Double = 0
                Dim SGST_TAXAMT As Double = 0
                Dim gst As Double = 0
                Dim gst_amt As Double = 0
                Dim net As Double
                Dim rnd As Integer
                Dim myflag As Boolean = False
                Dim amtt As Double = 0
                If chkrs4.EOF = False Then
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                    Else
                        census = chkrs4.Fields(5).Value
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                    Else
                        survey = chkrs4.Fields(4).Value
                    End If
                    If Not IsDBNull(chkrs4.Fields(22).Value) Then
                        myflag = True
                    End If
                    pname = chkrs4.Fields(38).Value
                    pcode1 = chkrs4.Fields(1).Value
                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox1.Text & "' and GODWN_NO='" & ComboBox2.Text & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    If chkrs2.EOF = False Then
                        chkrs2.MoveFirst()
                        amtt = chkrs2.Fields(4).Value
                        If IsDBNull(chkrs2.Fields(5).Value) Then
                        Else
                            amtt = amtt + chkrs2.Fields(5).Value
                        End If
                    End If
                    chkrs2.Close()
                    chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs4.Fields(37).Value & "'", xcon)

                    If chkrs3.EOF = False Then
                        If IsDBNull(chkrs3.Fields(2).Value) Then
                            CGST_RATE = 0
                        Else
                            CGST_RATE = chkrs3.Fields(2).Value
                        End If
                        If IsDBNull(chkrs3.Fields(3).Value) Then
                            SGST_RATE = 0
                        Else
                            SGST_RATE = chkrs3.Fields(3).Value
                        End If
                    End If
                    gst = CGST_RATE + SGST_RATE
                    chkrs3.Close()
                    gst_amt = gst * amtt / 100
                    rnd = gst_amt - Math.Round(gst_amt)
                    If rnd >= 50 Then
                        gst_amt = Math.Round(gst_amt) + 1
                    Else
                        gst_amt = Math.Round(gst_amt)
                    End If

                    net = amtt + gst_amt
                    CGST_TAXAMT = gst_amt / 2
                    CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                End If
                chkrs4.Close()
                '''''''''''''''''''godown detail end''''''''''''''''''''''''''

                '''''''''''''''''''against bill and period start''''''''''''''
                Dim grp As String = ComboBox1.Text
                Dim gdn As String = ComboBox2.Text
                Dim invdt As DateTime = recDateList(X)
                Dim inv As Integer = recNoList(X)
                Dim FIRSTREC As Boolean = True
                Dim FROMNO As String = ""
                Dim TONO As String = ""
                Dim against As String = ""
                Dim against1 As String = ""
                Dim against3 As String = ""
                Dim against2 As String = ""
                Dim agcount As Integer = 0
                Dim adjusted_amt As Double = 0
                Dim last_bldate As DateTime
                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False

                    If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                        If FIRSTREC Then
                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs2.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs2.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                            If chkrs6.EOF = False Then
                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                            Else
                                FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                TONO = FROMNO
                            End If
                            chkrs6.Close()

                            FIRSTREC = False
                        Else
                            TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        End If

                        pname = chkrs2.Fields(15).Value
                        adjusted_amt = adjusted_amt + chkrs2.Fields(10).Value
                        last_bldate = chkrs2.Fields(4).Value
                        If agcount < 7 Then
                            against = against + "GO-" & chkrs2.Fields(0).Value & ", "
                        Else
                            If agcount < 14 Then
                                against1 = against1 + "GO-" & chkrs2.Fields(0).Value & ", "
                            Else
                                If agcount < 21 Then
                                    against2 = against2 + "GO-" & chkrs2.Fields(0).Value & ", "
                                Else
                                    If agcount < 28 Then
                                        against3 = against3 + "GO-" & chkrs2.Fields(0).Value & ", "
                                    Else
                                    End If
                                End If
                            End If
                        End If

                        agcount = agcount + 1


                    End If
                    If chkrs2.EOF = False Then
                        chkrs2.MoveNext()
                    End If
                    If chkrs2.EOF = True Then
                        Exit Do
                    End If
                Loop
                chkrs2.Close()
                If against.Length > 2 Then
                    against = against.Substring(0, against.Length - 2)
                End If
                If Trim(against1).Equals("") Then
                Else
                    against1 = against1.Substring(0, against1.Length - 2)
                End If
                If Trim(against2).Equals("") Then
                Else
                    against2 = against2.Substring(0, against2.Length - 2)
                End If
                If Trim(against3).Equals("") Then
                Else
                    against3 = against3.Substring(0, against3.Length - 2)
                End If
                '''''''''''''''''''against bill and period end''''''''''''

                ''''''''''''''find out if any advance is left after adjustment start
                Dim lastbilladjusted As Integer = 0
                Dim advanceamt As Double = 0
                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' and t2.[P_CODE] ='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs5.Open(Rss, xcon)
                    Do While chkrs5.EOF = False
                        If chkrs1.Fields(3).Value >= chkrs5.Fields(4).Value Then
                            lastbilladjusted = chkrs5.Fields(0).Value
                            last_bldate = chkrs5.Fields(4).Value
                        End If
                        If chkrs5.EOF = False Then
                            chkrs5.MoveNext()
                        End If
                    Loop
                    chkrs5.Close()
                    If lastbilladjusted = 0 Then
                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND [STATUS]='C' AND P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                        chkrs5.Open(Rss, xcon)
                        If chkrs5.EOF = False Then
                            last_bldate = chkrs5.Fields(0).Value
                        End If
                        chkrs5.Close()
                    End If
                    Dim dtcounter As Integer = 1
                    If against.Length >= 1 Or lastbilladjusted = 0 Then
                        dtcounter = 1
                    Else
                        dtcounter = 2
                    End If
                    Do Until advanceamt <= 0
                        Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                        If lastbilladjusted = 0 Then
                            sdt = Convert.ToDateTime(last_bldate)
                        End If
                        If FIRSTREC Then
                            If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                advanceamt = advanceamt - net
                                TONO = FROMNO
                                FIRSTREC = False
                            End If
                        Else
                            TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                            advanceamt = advanceamt - net
                            dtcounter = dtcounter + 1
                        End If
                    Loop
                End If
                '''''''''''''find out if any advance is left after adjustment end
                Print(fnum, GetStringToPrint(13, "Party Name : ", "S") & GetStringToPrint(55, pname, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 1
                Print(fnum, GetStringToPrint(13, "Godown No. : ", "S") & GetStringToPrint(15, Trim(chkrs1.Fields(1).Value) & " " & Trim(chkrs1.Fields(2).Value), "S") & vbNewLine)
                Print(fnum, GetStringToPrint(13, "Survey No. : ", "S") & GetStringToPrint(12, survey, "S") & "    " & GetStringToPrint(13, "Census No. : ", "S") & GetStringToPrint(12, census, "S") & vbNewLine)
                xline = xline + 3
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 1
                Print(fnum, GetStringToPrint(20, "For the period from ", "S") & GetStringToPrint(45, Trim(FROMNO) & " to " & Trim(TONO), "S") & vbNewLine)


                If Trim(against1).Equals("") And chkrs1.Fields(6).Value = True Then
                    If Trim(against).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, " Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, against & ", Advance", "S") & vbNewLine)
                    End If


                Else
                    Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, against, "S") & vbNewLine)
                End If

                xline = xline + 2
                If Trim(against1).Equals("") Then
                Else
                    If Trim(against2).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against1 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against1, "S") & vbNewLine)
                    End If

                    xline = xline + 1
                End If
                If Trim(against2).Equals("") Then
                Else
                    If Trim(against3).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against2 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against2, "S") & vbNewLine)
                    End If
                    xline = xline + 1
                End If
                If Trim(against3).Equals("") Then
                Else
                    If chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against3 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against3, "S") & vbNewLine)
                    End If

                    xline = xline + 1
                End If
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 1

                Print(fnum, GetStringToPrint(13, "Amount Rs. : ", "S") & GetStringToPrint(10, Format(chkrs1.Fields(5).Value, "#####0.00"), "S"))
                'xline = xline + 1
                If chkrs1.Fields(7).Value.Equals("C") Then
                    Print(fnum, GetStringToPrint(8, "By Cash ", "S") & vbNewLine)
                    xline = xline + 1
                Else
                    Print(fnum, GetStringToPrint(10, "By Cheque ", "S") & vbNewLine)
                    xline = xline + 1
                End If
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(chkrs1.Fields(5).Value)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(49, inwordd.Length - 49)
                    Print(fnum, GetStringToPrint(16, "In Words   : Rs.", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(15) & GetStringToPrint(50, inword1, "S") & vbNewLine)
                    xline = xline + 2
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(16, "In Words   : Rs.", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    xline = xline + 1
                End If

                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 1
                '''''''''''''''''rent detail
                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & pcode1 & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                Dim amt As Double
                Dim rnt As Double
                Dim prnt As Double
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()

                    amt = chkrs2.Fields(4).Value
                    rnt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amt = amt + chkrs2.Fields(5).Value
                        prnt = chkrs2.Fields(5).Value
                    End If
                End If
                chkrs2.Close()

                ''''''''''''''''''''''''''''''''''''''''''''sgst amount ,cgst amount, sgstnrate and cgst rate using rent, gst tables

                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & pcode1 & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                ''   Dim amtt As Double = 0
                amtt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()
                    amtt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amtt = amtt + chkrs2.Fields(5).Value
                    End If
                End If
                chkrs2.Close()
                chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & hsnm & "'", xcon)

                If chkrs3.EOF = False Then
                    If IsDBNull(chkrs3.Fields(2).Value) Then
                        CGST_RATE = 0
                    Else
                        CGST_RATE = chkrs3.Fields(2).Value
                    End If
                    If IsDBNull(chkrs3.Fields(3).Value) Then
                        SGST_RATE = 0
                    Else
                        SGST_RATE = chkrs3.Fields(3).Value
                    End If
                End If
                gst = CGST_RATE + SGST_RATE
                chkrs3.Close()
                gst_amt = gst * amtt / 100
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If

                net = amtt + gst_amt
                CGST_TAXAMT = gst_amt / 2


                'CGST_TAXAMT = amt * CGST_RATE / 100
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)

                ''''''''''''''''''''''''''''''''''''''''''''sgst amount ,cgst amount, sgstnrate and cgst rate using rent, gst tables
                '''''''''''''''''rent detail



                Print(fnum, GetStringToPrint(16, "Per Month  : Rs.", "S") & GetStringToPrint(9, Format(amt, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(CGST_TAXAMT, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(SGST_TAXAMT, "#####0.00"), "S") & vbNewLine)
                xline = xline + 1
                If chkrs1.Fields(7).Value.Equals("Q") Then
                    Print(fnum, GetStringToPrint(13, "Cheque No. : ", "S") & GetStringToPrint(15, chkrs1.Fields(10).Value, "S") & GetStringToPrint(11, "Bank Name", "S") & GetStringToPrint(35, chkrs1.Fields(8).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Branch     : ", "S") & GetStringToPrint(35, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    xline = xline + 2
                End If

                Dim mnth As Integer = chkrs1.Fields(5).Value / (amt + CGST_TAXAMT + SGST_TAXAMT)

                Print(fnum, GetStringToPrint(13, "Amt.Detail : ", "S") & GetStringToPrint(14, "Rent     : Rs.", "S") & GetStringToPrint(11, Format((amt * mnth), "0.00"), "N").Substring(6 - amt.ToString.Length) & GetStringToPrint(5, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(14, "CGST @" & CGST_RATE & "% : Rs.", "S") & GetStringToPrint(11, Format(CGST_TAXAMT * mnth, "0.00"), "N").Substring(6 - amt.ToString.Length) & vbNewLine)
                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(14, "SGST @" & SGST_RATE & "% : Rs.", "S") & GetStringToPrint(11, Format(CGST_TAXAMT * mnth, "0.00"), "N").Substring(6 - amt.ToString.Length) & vbNewLine)
                xline = xline + 3
                Print(fnum, GetStringToPrint(61, " ", "S") & StrDup(7, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(61, " ", "S") & StrDup(7, "-") & vbNewLine)
                xline = xline + 5
                If myflag Then
                    Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(21, "Authorised Signatory*", "S") & vbNewLine)
                Else
                    Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(20, "Authorised Signatory", "S") & vbNewLine)
                End If
                xline = xline + 1
                For cc As Integer = xline + 1 To 29
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                Next
                Print(fnum, StrDup(80, "-") & vbNewLine)
            Next
            chkrs1.Close()

        Next
        xcon.Close()
        FileClose(fnum)
        '''''load created .dat file in view form
        Form21.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprint1.dat", RichTextBoxStreamType.PlainText)
        Form21.Show()  '''''show report form
        CreatePDF(Application.StartupPath & "\Reports\Recprint1.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)    '''''convert .dat file to .pdf file
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        '''''convert .dat file to .pdf file
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 77

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.Orientation = PdfSharp.PageOrientation.Portrait
            pdfPage.TrimMargins.Left = 15

            pdfPage.Height = 852
            pdfPage.Width = 595

            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 10, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1

                line = readFile.ReadLine()

                If line Is Nothing Then
                    Exit While
                Else
                    If counter > 58 Then
                        counter = 1
                        pdfPage = pdf.AddPage()
                        graph = XGraphics.FromPdfPage(pdfPage)
                        font = New XFont("COURIER NEW", 10, XFontStyle.Regular)

                        pdfPage.TrimMargins.Left = 15

                        pdfPage.Width = 595    ' 842
                        pdfPage.Height = 852
                        yPoint = 77
                    End If
                    Dim image As XImage = image.FromFile(Application.StartupPath & "\logo.png")
                    If counter = 1 Then
                        If ChkLogo.Checked Then
                            graph.DrawImage(image, 0, 0, image.Width, image.Height)
                        End If
                        yPoint = image.Height + 25
                        ' If
                    Else
                        If counter = 30 Then
                            If ChkLogo.Checked Then
                                graph.DrawImage(image, 0, yPoint, image.Width, image.Height)
                            End If
                            yPoint = yPoint + image.Height - 20
                        Else
                            font = New XFont("COURIER NEW", 10, XFontStyle.Regular)
                        End If
                    End If

                    graph.DrawString(line, font, XBrushes.Black,
                    New XRect(30, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = invoice_no & ".pdf"
            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        '''''allow only numerics
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub FrmRecPrnGdnwise_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ''''if focus is on godown number combobox and user press F1 key Help form will be visible
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "ComboBox2") Then
            helpgrpcombo = ComboBox1
            helpgdncombo = ComboBox2
            GodownHelp.Show()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''''''report print
        If DataGridView1.RowCount <= 0 Then
            MsgBox("No data exist for this godown")
            ComboBox1.Focus()
            Exit Sub
        End If
        If (ComboBox2.Text = "") Then
            MsgBox("Please enter Godown number")
            ComboBox2.Focus()
            Exit Sub
        End If
        Dim recNoList As New List(Of String)()
        Dim recDateList As New List(Of Date)()
        Dim xline As Integer = 0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, X).Value = True Then
                recNoList.Add(GetValue(DataGridView1.Item(5, X).Value))
                recDateList.Add(Convert.ToDateTime(DataGridView1.Item(4, X).Value))
            End If
        Next

        myArray = recNoList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Receipt")
            ComboBox2.Focus()
            Exit Sub
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''

        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports\")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprint1.dat", OpenMode.Output)
        Dim numRec As Integer = 0
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        For X As Integer = 0 To myArray.Length - 1
            Dim strn As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE=Format('" & Convert.ToDateTime(recDateList(X)) & "','dd/mm/yyyy') AND REC_NO=" & recNoList(X) & " order by [RECEIPT].REC_NO"
            chkrs1.Open(strn, xcon)
            If chkrs1.BOF = False Then
                chkrs1.MoveFirst()
            End If
            For xX As Integer = 1 To Convert.ToInt16(TextBox6.Text.Trim)

                Print(fnum, GetStringToPrint(17, "Receipt No.: GST-", "S") & GetStringToPrint(5, Trim(recNoList(X)), "S") & GetStringToPrint(41, " ", "S") & GetStringToPrint(7, "Date : ", "S") & GetStringToPrint(10, Trim(recDateList(X)), "S") & vbNewLine)
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = 2
                ''''''''''''''''''''''godown detail start
                chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & ComboBox1.Text & "' AND GODWN_NO='" & ComboBox2.Text & "' AND [STATUS]='C' ", xcon)
                Dim census As String = ""
                Dim survey As String = ""
                Dim pname As String = ""
                Dim rent As Double = 0
                Dim pcode1 As String = ""
                Dim hsnm As String = ""
                Dim CGST_RATE As Double = 0
                Dim SGST_RATE As Double = 0
                Dim CGST_TAXAMT As Double = 0
                Dim SGST_TAXAMT As Double = 0
                Dim gst As Double = 0
                Dim gst_amt As Double = 0
                Dim net As Double
                Dim rnd As Integer
                Dim myflag As Boolean = False
                Dim amtt As Double = 0
                If chkrs4.EOF = False Then
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                    Else
                        census = chkrs4.Fields(5).Value
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                    Else
                        survey = chkrs4.Fields(4).Value
                    End If
                    If Not IsDBNull(chkrs4.Fields(22).Value) Then
                        myflag = True
                    End If
                    pname = chkrs4.Fields(38).Value
                    pcode1 = chkrs4.Fields(1).Value
                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox1.Text & "' and GODWN_NO='" & ComboBox2.Text & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    If chkrs2.EOF = False Then
                        chkrs2.MoveFirst()
                        amtt = chkrs2.Fields(4).Value
                        If IsDBNull(chkrs2.Fields(5).Value) Then
                        Else
                            amtt = amtt + chkrs2.Fields(5).Value
                        End If
                    End If
                    chkrs2.Close()
                    chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs4.Fields(37).Value & "'", xcon)

                    If chkrs3.EOF = False Then
                        If IsDBNull(chkrs3.Fields(2).Value) Then
                            CGST_RATE = 0
                        Else
                            CGST_RATE = chkrs3.Fields(2).Value
                        End If
                        If IsDBNull(chkrs3.Fields(3).Value) Then
                            SGST_RATE = 0
                        Else
                            SGST_RATE = chkrs3.Fields(3).Value
                        End If
                    End If
                    gst = CGST_RATE + SGST_RATE
                    chkrs3.Close()
                    gst_amt = gst * amtt / 100
                    rnd = gst_amt - Math.Round(gst_amt)
                    If rnd >= 50 Then
                        gst_amt = Math.Round(gst_amt) + 1
                    Else
                        gst_amt = Math.Round(gst_amt)
                    End If

                    net = amtt + gst_amt
                    CGST_TAXAMT = gst_amt / 2
                    CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                End If
                chkrs4.Close()

                '''''''''''''''''''godown detail end

                '''''''''''''''''''against bill and period start''''''''''''''
                Dim grp As String = ComboBox1.Text
                Dim gdn As String = ComboBox2.Text
                Dim invdt As DateTime = recDateList(X)
                Dim inv As Integer = recNoList(X)
                Dim FIRSTREC As Boolean = True
                Dim FROMNO As String = ""
                Dim TONO As String = ""
                Dim against As String = ""
                Dim against1 As String = ""
                Dim against3 As String = ""
                Dim against2 As String = ""
                Dim agcount As Integer = 0
                Dim adjusted_amt As Double = 0
                Dim last_bldate As DateTime
                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False

                    If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                        '''''''''''''''''''''''''''''get from month and to month for adjustment
                        If FIRSTREC Then
                            chkrs6.Open("Select FROM_DATE,TO_DATE FROM BILL_TR WHERE [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND INVOICE_NO='" & chkrs2.Fields(0).Value & "' and  BILL_DATE=format('" & Convert.ToDateTime(chkrs2.Fields(4).Value) & "','dd/mm/yyyy') ", xcon)
                            If chkrs6.EOF = False Then
                                FROMNO = MonthName(Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("FROM_DATE").Value).Year
                                TONO = MonthName(Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Month, False) & " - " & Convert.ToDateTime(chkrs6.Fields("TO_DATE").Value).Year
                            Else
                                FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                                TONO = FROMNO
                            End If
                            chkrs6.Close()

                            FIRSTREC = False
                        Else
                            TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & " - " & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        End If
                        '''''''''''''''''''''''''''''get from month and to month for adjustment

                        pname = chkrs2.Fields(15).Value
                        adjusted_amt = adjusted_amt + chkrs2.Fields(10).Value
                        last_bldate = chkrs2.Fields(4).Value
                        If agcount < 7 Then
                            against = against + "GO-" & chkrs2.Fields(0).Value & ", "
                        Else
                            If agcount < 14 Then
                                against1 = against1 + "GO-" & chkrs2.Fields(0).Value & ", "
                            Else
                                If agcount < 21 Then
                                    against2 = against2 + "GO-" & chkrs2.Fields(0).Value & ", "
                                Else
                                    If agcount < 28 Then
                                        against3 = against3 + "GO-" & chkrs2.Fields(0).Value & ", "
                                    Else
                                    End If
                                End If
                            End If
                        End If

                        agcount = agcount + 1

                    End If
                    If chkrs2.EOF = False Then
                        chkrs2.MoveNext()
                    End If
                    If chkrs2.EOF = True Then
                        Exit Do
                    End If
                Loop
                chkrs2.Close()
                If against.Length > 2 Then
                    against = against.Substring(0, against.Length - 2)
                End If
                If Trim(against1).Equals("") Then
                Else
                    against1 = against1.Substring(0, against1.Length - 2)
                End If
                If Trim(against2).Equals("") Then
                Else
                    against2 = against2.Substring(0, against2.Length - 2)
                End If
                If Trim(against3).Equals("") Then
                Else
                    against3 = against3.Substring(0, against3.Length - 2)
                End If
                '''''''''''''''''''against bill and period end''''''''''''

                ''''''''''''''find out if any advance is left after adjustment start
                Dim lastbilladjusted As Integer = 0
                Dim advanceamt As Double = 0
                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' and t2.[P_CODE] ='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs5.Open(Rss, xcon)
                    Do While chkrs5.EOF = False
                        If chkrs1.Fields(3).Value >= chkrs5.Fields(4).Value Then
                            lastbilladjusted = chkrs5.Fields(0).Value
                            last_bldate = chkrs5.Fields(4).Value
                        End If
                        If chkrs5.EOF = False Then
                            chkrs5.MoveNext()
                        End If
                    Loop
                    chkrs5.Close()
                    If lastbilladjusted = 0 Then
                        Rss = "SELECT FROM_D From GODOWN Where [GROUP] ='" & grp & "' AND GODWN_NO='" & gdn & "' AND [STATUS]='C' AND P_CODE='" & pcode1 & "' order by [GROUP]+GODWN_NO"
                        chkrs5.Open(Rss, xcon)
                        If chkrs5.EOF = False Then
                            last_bldate = chkrs5.Fields(0).Value
                        End If
                        chkrs5.Close()
                    End If
                    Dim dtcounter As Integer = 1
                    If against.Length >= 1 Or lastbilladjusted = 0 Then
                        dtcounter = 1
                    Else
                        dtcounter = 2
                    End If
                    Do Until advanceamt <= 0
                        Dim sdt As Date = Convert.ToDateTime(last_bldate).AddMonths(1)
                        If lastbilladjusted = 0 Then
                            sdt = Convert.ToDateTime(last_bldate)
                        End If
                        If FIRSTREC Then
                            If IsDBNull(FROMNO) Or FROMNO = Nothing Then
                                FROMNO = MonthName(sdt.Month, False) & "-" & sdt.Year
                                advanceamt = advanceamt - net
                                TONO = FROMNO
                                FIRSTREC = False
                            End If
                        Else
                            TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                            advanceamt = advanceamt - net
                            dtcounter = dtcounter + 1
                        End If
                    Loop
                End If
                '''''''''''''find out if any advance is left after adjustment end

                Print(fnum, GetStringToPrint(13, "Party Name : ", "S") & GetStringToPrint(55, pname, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 2
                Print(fnum, GetStringToPrint(13, "Godown No. : ", "S") & GetStringToPrint(15, Trim(chkrs1.Fields(1).Value) & " " & Trim(chkrs1.Fields(2).Value), "S") & vbNewLine)
                Print(fnum, GetStringToPrint(13, "Survey No. : ", "S") & GetStringToPrint(12, survey, "S") & "    " & GetStringToPrint(13, "Census No. : ", "S") & GetStringToPrint(12, census, "S") & vbNewLine)
                xline = xline + 2
                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                xline = xline + 1
                Print(fnum, GetStringToPrint(20, "For the period from ", "S") & GetStringToPrint(45, Trim(FROMNO) & " to " & Trim(TONO), "S") & vbNewLine)
                xline = xline + 1



                If Trim(against1).Equals("") And chkrs1.Fields(6).Value = True Then
                    If Trim(against).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, " Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, against & ", Advance", "S") & vbNewLine)
                    End If


                Else
                    Print(fnum, GetStringToPrint(13, "Against Bill : ", "S") & GetStringToPrint(68, against, "S") & vbNewLine)
                End If

                xline = xline + 1
                If Trim(against1).Equals("") Then
                Else
                    If Trim(against2).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against1 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against1, "S") & vbNewLine)
                    End If

                    xline = xline + 1
                End If
                If Trim(against2).Equals("") Then
                Else
                    If Trim(against3).Equals("") And chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against2 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against2, "S") & vbNewLine)
                    End If
                    xline = xline + 1
                End If
                If Trim(against3).Equals("") Then
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    xline = xline + 1
                Else
                    If chkrs1.Fields(6).Value = True Then
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against3 & ", Advance", "S") & vbNewLine)
                    Else
                        Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(63, against3, "S") & vbNewLine)
                    End If

                    xline = xline + 1
                End If
                Print(fnum, GetStringToPrint(13, "Amount Rs. : ", "S") & GetStringToPrint(10, Format(chkrs1.Fields(5).Value, "#####0.00"), "S"))

                If chkrs1.Fields(7).Value.Equals("C") Then
                    Print(fnum, GetStringToPrint(8, "By Cash ", "S") & vbNewLine)
                    xline = xline + 1
                Else
                    Print(fnum, GetStringToPrint(10, "By Cheque ", "S") & vbNewLine)
                    xline = xline + 1
                End If
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(chkrs1.Fields(5).Value)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(49, inwordd.Length - 49)
                    Print(fnum, GetStringToPrint(16, "In Words   : Rs.", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(15) & GetStringToPrint(50, inword1, "S") & vbNewLine)
                    xline = xline + 2
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(16, "In Words   : Rs.", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    xline = xline + 1
                End If
                If Trim(against2).Equals("") Then
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    xline = xline + 1
                Else

                End If
                '''''''''''''''''rent detail
                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & pcode1 & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                Dim amt As Double
                Dim rnt As Double
                Dim prnt As Double
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()

                    amt = chkrs2.Fields(4).Value
                    rnt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amt = amt + chkrs2.Fields(5).Value
                        prnt = chkrs2.Fields(5).Value
                    End If
                End If
                chkrs2.Close()

                ''''''''''''''''''''''''''''''''''''''''''''sgst amount ,cgst amount, sgstnrate and cgst rate using rent, gst tables

                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & pcode1 & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                ''   Dim amtt As Double = 0
                amtt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()
                    amtt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amtt = amtt + chkrs2.Fields(5).Value
                    End If
                End If
                chkrs2.Close()
                chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & hsnm & "'", xcon)

                If chkrs3.EOF = False Then
                    If IsDBNull(chkrs3.Fields(2).Value) Then
                        CGST_RATE = 0
                    Else
                        CGST_RATE = chkrs3.Fields(2).Value
                    End If
                    If IsDBNull(chkrs3.Fields(3).Value) Then
                        SGST_RATE = 0
                    Else
                        SGST_RATE = chkrs3.Fields(3).Value
                    End If
                End If
                gst = CGST_RATE + SGST_RATE
                chkrs3.Close()
                gst_amt = gst * amtt / 100
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If

                net = amtt + gst_amt
                CGST_TAXAMT = gst_amt / 2


                'CGST_TAXAMT = amt * CGST_RATE / 100
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)

                ''''''''''''''''''''''''''''''''''''''''''''sgst amount ,cgst amount, sgstnrate and cgst rate using rent, gst tables
                '''''''''''''''''rent detail


                Print(fnum, GetStringToPrint(16, "Per Month  : Rs.", "S") & GetStringToPrint(9, Format(amt, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(CGST_TAXAMT, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(SGST_TAXAMT, "#####0.00"), "S") & vbNewLine)
                xline = xline + 1
                If chkrs1.Fields(7).Value.Equals("Q") Then
                    Print(fnum, GetStringToPrint(13, "Cheque No. : ", "S") & GetStringToPrint(15, chkrs1.Fields(10).Value, "S") & GetStringToPrint(11, "Bank Name", "S") & GetStringToPrint(35, chkrs1.Fields(8).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Branch     : ", "S") & GetStringToPrint(35, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    xline = xline + 2
                End If

                Dim mnth As Integer = chkrs1.Fields(5).Value / (amt + CGST_TAXAMT + SGST_TAXAMT)

                Print(fnum, GetStringToPrint(13, "Amt.Detail : ", "S") & GetStringToPrint(14, "Rent     : Rs.", "S") & GetStringToPrint(11, Format((amt * mnth), "0.00"), "N").Substring(6 - amt.ToString.Length) & GetStringToPrint(5, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(14, "CGST @" & CGST_RATE & "% : Rs.", "S") & GetStringToPrint(11, Format(CGST_TAXAMT * mnth, "0.00"), "N").Substring(6 - amt.ToString.Length) & vbNewLine)
                Print(fnum, GetStringToPrint(13, " ", "S") & GetStringToPrint(14, "SGST @" & SGST_RATE & "% : Rs.", "S") & GetStringToPrint(11, Format(CGST_TAXAMT * mnth, "0.00"), "N").Substring(6 - amt.ToString.Length) & vbNewLine)
                xline = xline + 3
                Print(fnum, GetStringToPrint(61, " ", "S") & StrDup(7, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(60, " ", "S") & StrDup(1, "|") & StrDup(7, " ") & StrDup(1, "|") & vbNewLine)
                Print(fnum, GetStringToPrint(61, " ", "S") & StrDup(7, "-") & vbNewLine)
                xline = xline + 5
                If myflag Then
                    Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(21, "Authorised Signatory*", "S") & vbNewLine)
                Else
                    Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(20, "Authorised Signatory", "S") & vbNewLine)
                End If
                xline = xline + 1
                For cc As Integer = xline + 1 To 29
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                Next
            Next
            chkrs1.Close()
        Next
        xcon.Close()
        FileClose(fnum)
        ''''display created .dat file in report view form
        Form21.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprint1.dat", RichTextBoxStreamType.PlainText)
        Form21.Show()   ''''show report form
        CreatePDF(Application.StartupPath & "\Reports\Recprint1.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)  '''''create pdf file for the .dat file
        ''''''''send pdf file to default printer
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True
        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        '''''if user select godown number by typing characters in godown combo, fill receipt datagrid for selection criteria
        If godownfilled Then
            showdata(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub


End Class