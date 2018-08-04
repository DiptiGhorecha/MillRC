Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmSpReport
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim chkrs11 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim chkrs6 As New ADODB.Recordset
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim fnumm As Integer
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim formloaded As Boolean = False
    Public dBaseConnection As New System.Data.OleDb.OleDbConnection
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DaysInMonth As Integer = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim MMT As Int16 = Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim MMT1 As String
        If MMT < 10 Then
            MMT1 = "0" + MMT.ToString()
        End If
        Dim startPP As Date = New Date(Convert.ToInt16(ComboBox4.Text), MMT1, DaysInMonth)
        Dim stdt As String = "SELECT sUM(NET_AMOUNT) AS NET FROM BILL where MONTH(BILL_DATE)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR(BILL_DATE)='" & ComboBox4.Text & "' AND HSN='997212'"
        Dim objPRNSetup = New clsPrinterSetup
        'set Paper Lines and Left Margin
        prnmaxpagelines = objPRNSetup.LinesPerPage
        If objPRNSetup.PageSize = PRNA4Paper Then
            prnleftmargin = 7
        Else
            prnleftmargin = 2
        End If
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        fnumm = 2 'FreeFile() '''''''''Get FreeFile No. to wsrite to csv file '''''''''''
        xcount = 0      '''''''''Set xcount'''''''''''''''''
        xlimit = 88     '''''''''Set xlimit'''''''''''''''''
        Dim srno As Integer = 0
        'xpage = 1
        xpage = Val("2")

        Dim numRec As Integer = 0
        If (Not System.IO.Directory.Exists(Application.StartupPath & "\Reports")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Reports")
        End If
        FileOpen(fnum, Application.StartupPath & "\Reports\spreport.dat", OpenMode.Output)

        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        chkrs1.Open("SELECT SUM(BILL_AMOUNT) AS NET,SUM(CGST_AMT) AS CGST,SUM(SGST_AMT) AS SGST,SUM(NET_AMOUNT) AS NTOT FROM BILL where MONTH(BILL_DATE)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR(BILL_DATE)='" & ComboBox4.Text & "' AND HSN='997212'", xcon)
        Dim HSN As String
        Dim HSNPRINT As Boolean
        Dim totnet, tottaxable, totcgst, totsgst, groupnet, grouptaxable, groupcgst, groupsgst As Double
        Dim atotnet, atottaxable, atotcgst, atotsgst, agroupnet, agrouptaxable, agroupcgst, agroupsgst As Double
        Do While chkrs1.EOF = False
            If chkrs1.Fields(0).Value.Equals("") Or IsDBNull(chkrs1.Fields(0).Value) Then

                Print(fnum, GetStringToPrint(50, "Total Rent Bills less Residential ", "S") & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " = " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & vbNewLine)

            Else

                Print(fnum, GetStringToPrint(50, "Total Rent Bills less Residential ", "S") & GetStringToPrint(14, Format(chkrs1.Fields(0).Value, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(chkrs1.Fields(1).Value, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(chkrs1.Fields(2).Value, "##########0.00"), "N") & " = " & GetStringToPrint(14, Format(chkrs1.Fields(3).Value, "##########0.00"), "N") & vbNewLine)

            End If
            If (chkrs1.EOF = False) Then
                chkrs1.MoveNext()
            End If
        Loop
        chkrs1.Close()
        chkrs1.Open("SELECT SUM(BILL_AMOUNT) AS NET,SUM(CGST_AMT) AS CGST,SUM(SGST_AMT) AS SGST,SUM(NET_AMOUNT) AS NTOT FROM BILL where MONTH(BILL_DATE)='" & DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month & "' AND YEAR(BILL_DATE)='" & ComboBox4.Text & "' AND HSN='997212' AND ADVANCE=TRUE", xcon)
        Do While chkrs1.EOF = False
            If chkrs1.Fields(0).Value.Equals("") Or IsDBNull(chkrs1.Fields(0).Value) Then
                Print(fnum, GetStringToPrint(50, "Adj against Advance less Residential ", "S") & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & " = " & GetStringToPrint(14, Format(0, "##########0.00"), "N") & vbNewLine)
            Else
                Print(fnum, GetStringToPrint(50, "Adj against Advance less Residential ", "S") & GetStringToPrint(14, Format(chkrs1.Fields(0).Value, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(chkrs1.Fields(1).Value, "##########0.00"), "N") & " + " & GetStringToPrint(14, Format(chkrs1.Fields(2).Value, "##########0.00"), "N") & " = " & GetStringToPrint(14, Format(chkrs1.Fields(3).Value, "##########0.00"), "N") & vbNewLine)
            End If
            If (chkrs1.EOF = False) Then
                chkrs1.MoveNext()
            End If
        Loop
        chkrs1.Close()

        DaysInMonth = 1  'Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        '  MsgBox(DaysInMonth)
        Dim startP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        DaysInMonth = Date.DaysInMonth(ComboBox4.Text, DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month)
        Dim endP As DateTime = New DateTime(Convert.ToInt16(ComboBox4.Text), Convert.ToInt16(DateTime.ParseExact(ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month), DaysInMonth)
        chkrs11.Open("SELECT * FROM GST where hsn_no='997212' ORDER BY HSN_NO", xcon)
        If chkrs11.BOF = False Then
            chkrs11.MoveFirst()
        End If

        Dim first As Boolean = True
        Dim totamt As Double = 0
        Dim totadv As Double = 0
        Do While chkrs11.EOF = False
            HSN = chkrs11.Fields(0).Value
            HSNPRINT = True
            'If first Then

            '    Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
            '    Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
            '    Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
            '    Print(fnum, StrDup(180, "=") & vbNewLine)
            '    first = False
            '    xline = xline + 3
            'End If
            ' Print(fnum, GetStringToPrint(35, "HSN Number :" & chkrs11.Fields(0).Value, "S") & GetStringToPrint(75, chkrs11.Fields(1).Value, "S") & vbNewLine)
            HSNPRINT = False
            ' xline = xline + 1


            Dim str As String = "SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO"

            chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
            'chkrs1.Open("SELECT [RECEIPT].* from [RECEIPT] where [RECEIPT].REC_DATE between FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by [RECEIPT].REC_DATE+[RECEIPT].REC_NO", xcon)
            If chkrs1.BOF = False Then
                chkrs1.MoveFirst()
            End If
            Dim gtotamt As Double = 0
            Dim gtotadv As Double = 0
            Do While chkrs1.EOF = False
                '  If chkrs1.Fields(4).Value >= strrec And chkrs1.Fields(4).Value <= edrec Then
                'If first Then

                '            Print(fnum, GetStringToPrint(17, "Receipt Date", "S") & GetStringToPrint(17, "Receipt No.", "S") & GetStringToPrint(13, "Amount", "N") & GetStringToPrint(13, "  Advance", "S") & GetStringToPrint(12, "Cash/Cheque", "S") & GetStringToPrint(7, "Group", "S") & GetStringToPrint(13, "Godown No.", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(33, "Bank A/C Detail", "S") & vbNewLine)
                '            Print(fnum, GetStringToPrint(17, "Cheque No.", "S") & GetStringToPrint(17, "Cheque Date", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, "Bank Name", "S") & GetStringToPrint(33, "Branch", "S") & vbNewLine)
                '            Print(fnum, GetStringToPrint(17, "From Month-Year", "S") & GetStringToPrint(17, "To Month-Year", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, "Adjusted Bill No.", "S") & vbNewLine)
                '            Print(fnum, StrDup(180, "=") & vbNewLine)
                '            first = False
                '            xline = xline + 3
                '        End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                chkrs4.Open("SELECT [GODOWN].*,[PARTY].P_NAME from [GODOWN] INNER JOIN [PARTY] on [GODOWN].P_CODE=[PARTY].P_CODE WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' AND GODWN_NO='" & chkrs1.Fields(2).Value & "' AND [STATUS]='C'", xcon)
                Dim HSNCD As String = ""
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
                    HSNCD = chkrs4.Fields(37).Value

                    chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs1.Fields(1).Value & "' and GODWN_NO='" & chkrs1.Fields(2).Value & "' and P_CODE ='" & chkrs4.Fields(1).Value & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
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

                Dim grp As String = chkrs1.Fields(1).Value
                Dim gdn As String = chkrs1.Fields(2).Value
                Dim invdt As DateTime = chkrs1.Fields(3).Value
                Dim inv As Integer = chkrs1.Fields(4).Value
                Dim FIRSTREC As Boolean = True
                Dim FROMNO As String = ""
                Dim TONO As String
                Dim against As String = ""
                Dim against1 As String = ""
                Dim against3 As String = ""
                Dim against2 As String = ""
                Dim agcount As Integer = 0
                Dim adjusted_amt As Double = 0
                Dim last_bldate As DateTime
                '  If chkrs1.Fields(6).Value = True Then

                '  Else
                '  grp = chkrs1.Fields(1).Value
                '   gdn = chkrs1.Fields(2).Value
                Dim RS As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=#" & Convert.ToDateTime(invdt) & "#)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
                chkrs2.Open("SELECT t2.*,[PARTY].P_NAME,(SELECT SUM(NET_AMOUNT) FROM [BILL] as t1 Where t1.[GROUP] ='" & grp & "' AND t1.GODWN_NO='" & gdn & "' AND (t1.REC_NO='" & inv & "' and  t1.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) AS balance,IIF(t2.rec_no Is Not null,TRUE,FALSE) AS checker From [BILL] As t2 INNER Join [PARTY] On t2.P_CODE=[PARTY].P_CODE Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND ((t2.REC_NO='" & inv & "' AND t2.REC_DATE=format('" & Convert.ToDateTime(invdt) & "','dd/mm/yyyy'))) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO", xcon)

                Do While chkrs2.EOF = False
                    'sgsrate = chkrs2.Fields(8).Value
                    'cgsrate = chkrs2.Fields(6).Value
                    'sgamt = chkrs2.Fields(9).Value
                    'cgamt = chkrs2.Fields(7).Value
                    HSNCD = chkrs2.Fields(11).Value
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
                            TONO = MonthName(Convert.ToDateTime(chkrs2.Fields(4).Value).Month, False) & "-" & Convert.ToDateTime(chkrs2.Fields(4).Value).Year
                        End If
                        last_bldate = chkrs2.Fields(4).Value
                        pname = chkrs2.Fields(16).Value
                        adjusted_amt = adjusted_amt + chkrs2.Fields(10).Value
                        If agcount < 8 Then
                            against = against + "GO-" & chkrs2.Fields(0).Value & ", "
                        Else
                            If agcount < 15 Then
                                against1 = against1 + "GO-" & chkrs2.Fields(0).Value & ", "
                            Else
                                If agcount < 22 Then
                                    against2 = against2 + "GO-" & chkrs2.Fields(0).Value & ", "
                                Else
                                    If agcount < 29 Then
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
                '    Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)

                ''''''''''''''find out if any advance is left after adjustment start
                Dim advanceamt As Double = 0
                Dim advanceamtprint As Double = 0
                Dim lastbilladjusted As Integer

                advanceamt = chkrs1.Fields(5).Value - adjusted_amt
                advanceamtprint = advanceamt
                If advanceamt > 0 Then
                    Dim Rss As String = "SELECT T2.INVOICE_NO,T2.GROUP,T2.GODWN_NO,T2.P_CODE,T2.BILL_DATE,T2.BILL_AMOUNT,T2.CGST_RATE,T2.CGST_AMT,T2.SGST_RATE,T2.SGST_AMT,T2.NET_AMOUNT,T2.HSN,T2.SRNO,T2.REC_NO,T2.REC_DATE,[GODOWN].FROM_D From [BILL] As t2 inner join GODOWN on t2.[GROUP]=[GODOWN].[GROUP] AND t2.[GODWN_NO]=[GODOWN].GODWN_NO Where t2.[GROUP] ='" & grp & "' AND t2.GODWN_NO='" & gdn & "' AND T2.P_CODE='" & pcode1 & "' AND ((t2.REC_NO IS NOT NULL AND t2.REC_DATE IS NOT NULL)) order by t2.BILL_DATE,t2.GROUP,t2.GODWN_NO"
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
                            TONO = MonthName(last_bldate.AddMonths(dtcounter).Month, False) & "-" & last_bldate.AddMonths(dtcounter).Year
                            advanceamt = advanceamt - net
                            dtcounter = dtcounter + 1
                        End If
                    Loop
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If HSNCD.Equals(HSN) Then
                    totamt = totamt + chkrs1.Fields(5).Value
                    totadv = totadv + advanceamtprint
                    gtotamt = gtotamt + chkrs1.Fields(5).Value
                    gtotadv = gtotadv + advanceamtprint
                    '        If CheckBox1.Checked Then
                    '            If chkrs1.Fields(6).Value = True Then
                    '                Print(fnum, GetStringToPrint(17, chkrs1.Fields(3).Value, "S") & GetStringToPrint(17, "GST-" + chkrs1.Fields(4).Value.ToString, "S") & GetStringToPrint(13, Format(chkrs1.Fields(5).Value, "######0.00"), "N") & GetStringToPrint(13, Format(advanceamtprint, "######0.00"), "N") & GetStringToPrint(12, "  " & chkrs1.Fields(7).Value, "S") & GetStringToPrint(7, chkrs1.Fields(1).Value, "S") & GetStringToPrint(13, chkrs1.Fields(2).Value, "S") & GetStringToPrint(50, pname, "S") & GetStringToPrint(33, IIf(IsDBNull(chkrs1.Fields(12).Value), "", chkrs1.Fields(12).Value), "S") & vbNewLine)

                    '                xline = xline + 1
                    '                Print(fnum, GetStringToPrint(17, chkrs1.Fields(10).Value, "S") & GetStringToPrint(17, IIf(chkrs1.Fields(7).Value.Equals("C"), "", "  " & chkrs1.Fields(11).Value), "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(50, chkrs1.Fields(8).Value, "S") & GetStringToPrint(33, chkrs1.Fields(9).Value, "S") & vbNewLine)
                    '                xline = xline + 1

                    '                Print(fnum, GetStringToPrint(17, FROMNO, "S") & GetStringToPrint(17, TONO, "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against, "S") & vbNewLine)
                    '                xline = xline + 1
                    '                If Trim(against1).Equals("") Then
                    '                Else
                    '                    '    against1 = against1.Substring(0, against1.Length - 2)
                    '                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against1, "S") & vbNewLine)
                    '                    xline = xline + 1
                    '                End If
                    '                If Trim(against2).Equals("") Then
                    '                Else
                    '                    '   against2 = against2.Substring(0, against2.Length - 2)
                    '                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against2, "S") & vbNewLine)
                    '                    xline = xline + 1
                    '                End If
                    '                If Trim(against3).Equals("") Then
                    '                Else
                    '                    '  against3 = against3.Substring(0, against3.Length - 2)
                    '                    Print(fnum, GetStringToPrint(17, "", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, " ", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(12, "", "S") & GetStringToPrint(7, "", "N") & GetStringToPrint(13, "", "N") & GetStringToPrint(63, against3, "S") & vbNewLine)
                    '                    xline = xline + 1
                    '                End If
                    '                Print(fnum, GetStringToPrint(15, " ", "S") & vbNewLine)
                    '                xline = xline + 1
                End If

                'End If
                '    Print(fnum, StrDup(80, "-") & vbNewLine)
                '  Next
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                If chkrs1.EOF = True Then
                    Exit Do
                End If
                '  End If
            Loop
            'Print(fnum, StrDup(180, "-") & vbNewLine)
            'Print(fnum, GetStringToPrint(17, "Group Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(gtotamt, "########0.00"), "N") & GetStringToPrint(13, Format(gtotadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
            'Print(fnum, StrDup(180, " ") & vbNewLine)
            chkrs1.Close()




            If chkrs11.EOF = False Then
                chkrs11.MoveNext()
            End If
        Loop
        'Print(fnum, StrDup(180, " ") & vbNewLine)
        '  Print(fnum, StrDup(180, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(50, "Total Advances less Residential (Net)", "S") & GetStringToPrint(14, Format(totadv, "##########0.00"), "N") & vbNewLine)
        ' Print(fnum, GetStringToPrint(17, "Total-->", "S") & GetStringToPrint(17, "", "S") & GetStringToPrint(13, Format(totamt, "########0.00"), "N") & GetStringToPrint(13, Format(totadv, "######0.00"), "N") & GetStringToPrint(12, "  ", "S") & GetStringToPrint(7, "", "S") & GetStringToPrint(13, "", "S") & GetStringToPrint(50, "", "S") & GetStringToPrint(33, "", "S") & vbNewLine)
        chkrs11.Close()

        Dim tb As DataTable
        'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
        '  tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
        tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by BILLWR.BILL_DATE;")
        DataGridView1.DataSource = tb
        '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE

        groupnet = 0
        grouptaxable = 0
        Dim taxrate As Double = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            '   Print(fnum, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & GetStringToPrint(35, "WF-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(7, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(4, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(12, X).Value, "S") & vbNewLine)

            groupnet = groupnet + DataGridView1.Item(7, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(4, X).Value
        Next


        ' Print(fnum, StrDup(185, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(50, "Total Warehouse ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & vbNewLine)


        'tb = getdbasetable("select BILLWR.*,ACCMST.* from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' order by BILLWR.BILL_DATE;")
        tb = getdbasetable("select BILL.*,ACCMST.* from BILL INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILL.PARTY_COD WHERE BILL.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') order by BILL.BILL_DATE;")     '(ACCMST.GST_NO<>'' OR ACCMST.GST_NO='Null') ;")
        DataGridView1.DataSource = tb
        '''select BILLWR.PARTY_COD,GODWN_NO,BILL_NO,BILL_DATE,AMOUNT,TOTAL,TAX,NET,OUTSTAND,CGST_AMT,SGST_AMT,ACCMST.PARTY_COD,A_NM,A_ADD1,A_ADD2,A_ADD3,A_CITY,LED_FOL,GST_NO,EMAIL_ID from BILLWR INNER JOIN ACCMST ON ACCMST.PARTY_COD=BILLWR.PARTY_COD WHERE BILLWR.BILL_DATE BETWEEN FORMAT('" & startP & "','DD/MM/YYYY') AND FORMAT('" & endP & "','DD/MM/YYYY') AND ACCMST.GST_NO<>'' group by BILLWR.BILL_NO order by BILLWR.BILL_DATE

        groupnet = 0
        grouptaxable = 0
        taxrate = 18.0
        For X As Integer = 0 To DataGridView1.RowCount - 1
            '   MsgBox(DataGridView1.Item(18, X).Value)
            ' Print(fnum, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & GetStringToPrint(12, " 24-Gujarat", "S") & GetStringToPrint(7, "   N   ", "S") & GetStringToPrint(12, " Regular", "S") & GetStringToPrint(10, "         ", "S") & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & GetStringToPrint(15, " ", "S") & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            ' Print(fnumm, GetStringToPrint(16, DataGridView1.Item(18, X).Value, "S") & "," & GetStringToPrint(35, "WV-" & DataGridView1.Item(2, X).Value, "S") & "," & GetStringToPrint(13, DataGridView1.Item(3, X).Value, "S") & "," & GetStringToPrint(10, Format(DataGridView1.Item(27, X).Value, "######0.00"), "N") & "," & GetStringToPrint(12, " 24-Gujarat", "S") & "," & GetStringToPrint(7, "   N   ", "S") & "," & GetStringToPrint(12, " Regular", "S") & "," & GetStringToPrint(10, "         ", "S") & "," & GetStringToPrint(7, Format(taxrate, "###0.00"), "N") & "," & GetStringToPrint(14, Format(DataGridView1.Item(19, X).Value, "##########0.00"), "N") & "," & GetStringToPrint(15, " ", "S") & "," & GetStringToPrint(55, DataGridView1.Item(32, X).Value, "S") & vbNewLine)
            groupnet = groupnet + DataGridView1.Item(27, X).Value
            grouptaxable = grouptaxable + DataGridView1.Item(19, X).Value
        Next
        'Print(fnumm, " " & vbNewLine)
        ' Print(fnumm, " " & vbNewLine)
        '  Print(fnum, StrDup(185, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(50, "Total Warehouse rent (Gandalal) ", "S") & GetStringToPrint(14, Format(grouptaxable, "##########0.00"), "N") & vbNewLine)


        xcon.Close()

        FileClose(fnum)
        ''''''''''''''''''''''''hsnwise
        Form24.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\spreport.dat", RichTextBoxStreamType.PlainText)
        Form24.Show()
        CreatePDF(Application.StartupPath & "\Reports\spreport.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
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
    Public Function getdbasetable(ByVal SqlString As String) As DataTable
        Dim ReturnableTable As New DataTable
        Try
            OpendBConnection()
            Dim SelectCommand As New System.Data.OleDb.OleDbCommand(SqlString, dBaseConnection)
            Dim TableAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
            TableAdapter.SelectCommand = SelectCommand
            TableAdapter.Fill(ReturnableTable)
            Return ReturnableTable
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & SqlString, 16, "Error")
            End
        End Try
        Return ReturnableTable
    End Function

    Public Sub OpendBConnection()
        Try
            Dim ConnectionString As String
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\WHBILL; Extended Properties=dBase IV;"
            dBaseConnection = New System.Data.OleDb.OleDbConnection(ConnectionString)

            If dBaseConnection.State = 0 Then dBaseConnection.Open()
        Catch ex As Exception
            MsgBox(ex.Message, 16, "Error")
        End Try
    End Sub

    Private Sub FrmSpReport_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True

        If DateTime.Now.Month = 1 Then
            ComboBox3.Text = DateAndTime.MonthName(12)
            ComboBox4.Text = DateTime.Now.Year - 1
        Else
            ComboBox3.Text = DateAndTime.MonthName(DateTime.Now.Month - 1)
            ComboBox4.Text = DateTime.Now.Year
        End If
        formloaded = True
    End Sub

    Private Sub FrmSpReport_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class