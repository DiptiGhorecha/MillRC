Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
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
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        ' GroupBox5.Visible = False
        '  DataGridView2.Enabled = False
        For Each column As DataGridViewColumn In DataGridView1.Columns
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
        showdata(ComboBox1.Text, ComboBox2.Text)
    End Sub
    Public Function showdata(grp As String, gdn As String)
        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            If (ComboBox2.Text.Equals("")) Then
                da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].[GROUP]='" & grp & "' order by [RECEIPT].REC_NO", MyConn)
            Else
                da = New OleDb.OleDbDataAdapter("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].[GROUP]='" & grp & "' AND [RECEIPT].GODWN_NO='" & gdn & "' order by [RECEIPT].REC_NO", MyConn)
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
            showdata(ComboBox1.Text, "")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If godownfilled Then
            showdata(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub
    Function GetValue(Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        '  DataGridView2(e.ColumnIndex, e.RowIndex).[ReadOnly] = True
        If (DataGridView1.CurrentCell.ColumnIndex = 0) Then
            '  DataGridView2(e.ColumnIndex, e.RowIndex).[ReadOnly] = False
            'DataGridView2.BeginEdit(False)

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
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim recNoList As New List(Of String)()
        Dim recDateList As New List(Of Date)()
        Dim xline As Integer = 0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, X).Value = True Then
                recNoList.Add(GetValue(DataGridView1.Item(5, X).Value))
                recDateList.Add(Convert.ToDateTime(DataGridView1.Item(4, X).Value))

                ' FILE_NO = FILE_NO.Replace(" ", "_")
            End If
        Next
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''

        'Dim xline As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports\")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprint.dat", OpenMode.Output)
        myArray = recNoList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Receipt")
            ComboBox2.Focus()
            Exit Sub
        End If

        Dim numRec As Integer = 0
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        For X As Integer = 0 To myArray.Length - 1
            'Dim strn As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE=Format(#" & Convert.ToDateTime(recDateList(X)) & "#,'dd/mm/yyyy') AND REC_NO=" & recNoList(X) & " order by [RECEIPT].REC_NO"
            Dim strn As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE=Format('" & Convert.ToDateTime(recDateList(X)) & "','dd/mm/yyyy') AND REC_NO=" & recNoList(X) & " order by [RECEIPT].REC_NO"
            chkrs1.Open(strn, xcon)
            If chkrs1.EOF = False Then
                chkrs1.MoveFirst()
            End If
            Dim advance As Double = chkrs1.Fields(5).Value
            For xX As Integer = 1 To Convert.ToInt16(TextBox6.Text.Trim)

                Print(fnum, GetStringToPrint(6, " ", "S") & GetStringToPrint(50, "THE MOTILAL HIRABHAI ESTATE & WAREHOUSE LIMITED", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(14, " ", "S") & GetStringToPrint(60, "OUTSIDE PREM DARWAJA, AHMEDABAD - 380 002, Gujarat, INDIA.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(17, " ", "S") & GetStringToPrint(50, "Ph. :(079)22161537 E-mail: contact@mhwarehouse.com", "S") & vbNewLine)
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
                If chkrs4.EOF = False Then
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                    Else
                        census = chkrs4.Fields(5).Value
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                    Else
                        survey = chkrs4.Fields(4).Value
                    End If
                    pname = chkrs4.Fields(38).Value
                    pcode1 = chkrs4.Fields(1).Value
                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox1.Text & "' and GODWN_NO='" & ComboBox2.Text & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    Dim amtt As Double = 0
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


                    'CGST_TAXAMT = amt * CGST_RATE / 100
                    CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    'SGST_TAXAMT = amt * SGST_RATE / 100
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
                '  If chkrs1.Fields(6).Value = True Then

                '  Else

                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format(#" & Convert.ToDateTime(invdt) & "#,'dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False

                    If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                        'If FIRSTREC Then
                        '    FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        '    TONO = FROMNO
                        '    FIRSTREC = False
                        'Else
                        '    TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        'End If

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
                        pcode1 = chkrs2.Fields(3).Value
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
                '   End If
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

                '  Dim last_bldate As DateTime
                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' and t2.[P_CODE] ='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs5.Open(Rss, xcon)
                    Do While chkrs5.EOF = False
                        '  chkrs5.MoveLast()
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
                            ' chkrs5.MoveLast()
                            'lastbilladjusted = chkrs5.Fields(13).Value
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
                                'last_bldate = chkrs5.Fields(0).Value
                            End If
                        Else
                            '  TONO = MonthName(sdt.AddMonths(dtcounter).Month, False) & "-" & sdt.AddMonths(dtcounter).Year
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
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If
                '  Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                ' xline = xline + 1
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



                '''''''''''''''''rent detail
                Print(fnum, GetStringToPrint(16, "Per Month  : Rs.", "S") & GetStringToPrint(9, Format(amt, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(CGST_TAXAMT, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(SGST_TAXAMT, "#####0.00"), "S") & vbNewLine)
                xline = xline + 1
                ' Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                ' xline = xline + 1
                If chkrs1.Fields(7).Value.Equals("Q") Then
                    Print(fnum, GetStringToPrint(13, "Cheque No. : ", "S") & GetStringToPrint(15, chkrs1.Fields(10).Value, "S") & GetStringToPrint(11, "Bank Name", "S") & GetStringToPrint(35, chkrs1.Fields(8).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Branch     : ", "S") & GetStringToPrint(35, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    xline = xline + 2
                    ' Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    ' xline = xline + 1
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
                Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(20, "Authorised Signatory", "S") & vbNewLine)
                xline = xline + 1
                'MsgBox(xline)
                For cc As Integer = xline + 1 To 29
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                Next
                Print(fnum, StrDup(80, "-") & vbNewLine)
            Next
            chkrs1.Close()

        Next
        xcon.Close()
        FileClose(fnum)
        Form21.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprint.dat", RichTextBoxStreamType.PlainText)
        Form21.Show()
        CreatePDF(Application.StartupPath & "\Reports\Recprint.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)


    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.Orientation = PdfSharp.PageOrientation.Portrait
            pdfPage.TrimMargins.Left = 15

            pdfPage.Height = 842
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
                    If counter > 60 Then
                        counter = 1
                        pdfPage = pdf.AddPage()
                        graph = XGraphics.FromPdfPage(pdfPage)
                        font = New XFont("COURIER NEW", 10, XFontStyle.Regular)

                        pdfPage.TrimMargins.Left = 15

                        pdfPage.Width = 595    ' 842
                        pdfPage.Height = 842
                        yPoint = 60
                    End If
                    If counter = 1 Or counter = 31 Then
                        font = New XFont("COURIER NEW", 14, XFontStyle.Bold)
                    Else
                        font = New XFont("COURIER NEW", 10, XFontStyle.Regular)
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
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub FrmRecPrnGdnwise_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 And (Me.ActiveControl.Name = "ComboBox2") Then
            helpgrpcombo = ComboBox1
        helpgdncombo = ComboBox2
            GodownHelp.Show()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
        ' Dim startBill1 As String = TextBox1.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox1.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox1.Text.Substring(12, 3)
        ' Dim endBill1 As String = TextBox2.Text.Substring(0, 4) & "_" & DateTime.ParseExact(TextBox2.Text.Substring(8, 3), "MMM", CultureInfo.CurrentCulture).Month & "_" & TextBox2.Text.Substring(12, 3)
        Dim recNoList As New List(Of String)()
        Dim recDateList As New List(Of Date)()
        Dim xline As Integer = 0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, X).Value = True Then
                recNoList.Add(GetValue(DataGridView1.Item(5, X).Value))
                recDateList.Add(Convert.ToDateTime(DataGridView1.Item(4, X).Value))

                ' FILE_NO = FILE_NO.Replace(" ", "_")
            End If
        Next
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''

        'Dim xline As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports\")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\Recprint.dat", OpenMode.Output)
        myArray = recNoList.ToArray()
        If (myArray.Length < 1) Then
            MsgBox("Please select Receipt")
            ComboBox2.Focus()
            Exit Sub
        End If

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

                Print(fnum, GetStringToPrint(6, " ", "S") & GetStringToPrint(50, "THE MOTILAL HIRABHAI ESTATE & WAREHOUSE LIMITED", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(14, " ", "S") & GetStringToPrint(60, "OUTSIDE PREM DARWAJA, AHMEDABAD - 380 002, Gujarat, INDIA.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(17, " ", "S") & GetStringToPrint(50, "Ph. :(079)22161537 E-mail: contact@mhwarehouse.com", "S") & vbNewLine)
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
                If chkrs4.EOF = False Then
                    If IsDBNull(chkrs4.Fields(5).Value) Then
                    Else
                        census = chkrs4.Fields(5).Value
                    End If
                    If IsDBNull(chkrs4.Fields(4).Value) Then
                    Else
                        survey = chkrs4.Fields(4).Value
                    End If
                    pname = chkrs4.Fields(38).Value
                    pcode1 = chkrs4.Fields(1).Value
                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & ComboBox1.Text & "' and GODWN_NO='" & ComboBox2.Text & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                    Dim amtt As Double = 0
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


                    'CGST_TAXAMT = amt * CGST_RATE / 100
                    CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                    'SGST_TAXAMT = amt * SGST_RATE / 100
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
                '  If chkrs1.Fields(6).Value = True Then

                '  Else

                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False

                    If chkrs2.Fields(13).Value >= inv And chkrs2.Fields(14).Value <= invdt And chkrs1.Fields(3).Value >= chkrs2.Fields(4).Value Then
                        'If FIRSTREC Then
                        '    FROMNO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        '    TONO = FROMNO
                        '    FIRSTREC = False
                        'Else
                        '    TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        'End If
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
                        last_bldate = chkrs2.Fields(4).Value
                        pname = chkrs2.Fields(15).Value
                        adjusted_amt = adjusted_amt + chkrs2.Fields(10).Value
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
                '   End If
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

                ' Dim last_bldate As DateTime
                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' and t2.[P_CODE] ='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                    chkrs5.Open(Rss, xcon)
                    Do While chkrs5.EOF = False
                        '  chkrs5.MoveLast()
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
                            ' chkrs5.MoveLast()
                            'lastbilladjusted = chkrs5.Fields(13).Value
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
                                'last_bldate = chkrs5.Fields(0).Value
                            End If
                        Else
                            TONO = MonthName(sdt.AddMonths(dtcounter).Month, False) & "-" & sdt.AddMonths(dtcounter).Year
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
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If
                '  Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                ' xline = xline + 1
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



                '''''''''''''''''rent detail
                Print(fnum, GetStringToPrint(16, "Per Month  : Rs.", "S") & GetStringToPrint(9, Format(amt, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(CGST_TAXAMT, "#####0.00"), "S") & GetStringToPrint(3, " + ", "S") & GetStringToPrint(9, Format(SGST_TAXAMT, "#####0.00"), "S") & vbNewLine)
                xline = xline + 1
                ' Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                ' xline = xline + 1
                If chkrs1.Fields(7).Value.Equals("Q") Then
                    Print(fnum, GetStringToPrint(13, "Cheque No. : ", "S") & GetStringToPrint(15, chkrs1.Fields(10).Value, "S") & GetStringToPrint(11, "Bank Name", "S") & GetStringToPrint(35, chkrs1.Fields(8).Value, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Branch     : ", "S") & GetStringToPrint(35, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    xline = xline + 2
                    ' Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    ' xline = xline + 1
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
                Print(fnum, GetStringToPrint(54, " ", "S") & GetStringToPrint(20, "Authorised Signatory", "S") & vbNewLine)
                xline = xline + 1
                'MsgBox(xline)
                For cc As Integer = xline + 1 To 29
                    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                Next
                Print(fnum, StrDup(80, "-") & vbNewLine)
            Next
            chkrs1.Close()

        Next
        xcon.Close()
        FileClose(fnum)
        Form21.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\Recprint.dat", RichTextBoxStreamType.PlainText)
        Form21.Show()
        CreatePDF(Application.StartupPath & "\Reports\Recprint.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True
        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If godownfilled Then

            showdata(ComboBox1.Text, ComboBox2.Text)
        End If
    End Sub
End Class