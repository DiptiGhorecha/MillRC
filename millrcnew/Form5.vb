Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.OleDb
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports System.Globalization
Imports PdfSharp.Pdf.Advanced

Public Class FrmInvoice
    Dim chkrs As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs11 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    '''''''' used to open a temparory Recordset
    Dim fnum As Integer                 '''''''' used to store freefile no.
    Dim xcount                          '''''''' used to store pagelines
    Dim xlimit                          '''''''' used to store page limits
    Dim xpage
    Dim pwidth As Integer
    Dim XLBL As String
    Dim Div As Double
    Dim ChkFlag As Boolean
    Public bChanged As Boolean
    Public bCustom As Boolean
    Dim CW As Integer = Me.Width ' Current Width
    Dim CH As Integer = Me.Height ' Current Height
    Dim IW As Integer = Me.Width ' Initial Width
    Dim IH As Integer = Me.Height ' Initial Height
    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream
    Dim NewData As Boolean
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"

    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim strReportFilePath As String
    Dim formloaded As Boolean = False
    Private Function checkpr()
        Try
            Dim objPRNSetup = New clsPrinterSetup
            'set Paper Lines and Left Margin
            prnmaxpagelines = objPRNSetup.LinesPerPage
            If objPRNSetup.PageSize = PRNA4Paper Then
                prnleftmargin = 7
            Else
                prnleftmargin = 2
            End If
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If

            Dim array() As String = {"AE", "AA", "AB", "AC"}
            Dim subsql As String
            subsql = "SELECT * FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "'"
            chkrs1.Open("SELECT * FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "'", xcon)

            If chkrs1.EOF = False Then
                MsgBox("Bills for date " + DateTimePicker1.Value.ToString + " are already generated ")
                chkrs1.Close()
                Exit Function
            End If
            chkrs1.Close()
            Dim iDate As String
            Dim fDate As DateTime
            Dim oDate As String
            Dim foDate As DateTime
            If (Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) >= 4) Then
                iDate = "30/04/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) + 1))
                foDate = Convert.ToDateTime(oDate)
            Else
                iDate = "30/04/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) - 1))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                foDate = Convert.ToDateTime(oDate)
            End If

            ' MsgBox(fDate.ToString)
            ' MsgBox(foDate.ToString)
            Dim strn As String = "Select * FROM BILL where [BILL].bill_date>=Format('" & fDate & "', 'dd/mm/yyyy') and [BILL].bill_date<=Format('" & foDate & "', 'dd/mm/yyyy') order by Year([BILL].bill_date)+[BILL].INVOICE_NO"
            '''''''''''''get last serial number
            '''''chkrs11.Open("Select * FROM BILL where Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by Year([BILL].bill_date)+[BILL].INVOICE_NO", xcon)
            chkrs11.Open(strn, xcon)
            Dim srno As Integer = 0
            Do While chkrs11.EOF = False
                srno = chkrs11.Fields(0).Value
                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop
            chkrs11.Close()


            '''''''''''''get last serial number

            fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
            xcount = 0      '''''''''Set xcount'''''''''''''''''
            xlimit = 88     '''''''''Set xlimit'''''''''''''''''

            'xpage = 1
            xpage = Val("2")
            Dim i1 As Double
            ''''''''''''''''' open a file sharereg.txt'''''''''''
            ' FileOpen(fnum, Application.StartupPath & "\Invoices\RecordSlipView.dat", OpenMode.Output)
            '  Call header()
            Dim numRec As Integer = 0
            Dim STRR As String = "SELECT * FROM GODOWN where [STATUS]='C' AND FROM_D <=#" & Convert.ToDateTime(DateTimePicker1.Value) & "# order by [GROUP]+GODWN_NO "
            chkrs.Open("SELECT * FROM GODOWN where [STATUS]='C' AND FROM_D <=FORMAT('" & Convert.ToDateTime(DateTimePicker1.Value) & "','DD/MM/YYYY') order by [GROUP]+GODWN_NO ", xcon)
            Do While chkrs.EOF = False
                numRec = numRec + 1
                srno = srno + 1
                Dim INVOICE_NO As String
                Dim FILE_NO As String
                Dim FILE_NOtmp As String
                'INVOICE_NO = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & chkrs.Fields(0).Value & chkrs.Fields(3).Value     'String.Format("{0:000}", srno)
                FILE_NOtmp = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", numRec)
                ' FILE_NOtmp = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & chkrs.Fields(0).Value & chkrs.Fields(3).Value     'String.Format("{0:000}", srno)
                'INVOICE_NO = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", srno)
                INVOICE_NO = String.Format("{0:0000}", srno)
                FILE_NO = INVOICE_NO.Replace("/", "_")   'DateTimePicker1.Value.Year & "_" & DateTimePicker1.Value.Month & "_" & chkrs.Fields(0).Value & chkrs.Fields(3).Value.replace("/", "_")
                FILE_NO = FILE_NO.Replace(" ", "_")
                ' FILE_NO = FILE_NO.Replace("-", "_")
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
                End If


                FileOpen(fnum, Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat", OpenMode.Output)
                Dim TENANT_CODE As String
                Dim TENANT_NAME As String
                Dim T_ADDREASS As String
                Dim TAD1 As String
                Dim TAD2 As String
                Dim TAD3 As String
                Dim TCITY As String
                Dim TSTATE As String
                Dim STATE_CODE As String
                Dim TEMAIL As String
                Dim TGST As String
                Dim amt As Double
                Dim CGST_TAXAMT As Double
                Dim SGST_TAXAMT As Double
                Dim CGST_RATE As Double
                Dim gst_amt As Double
                Dim SGST_RATE As Double
                Dim gst As Double
                TENANT_CODE = chkrs.Fields(1).Value
                chkrs1.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
                If chkrs1.EOF = False Then
                    TENANT_NAME = LTrim(chkrs1.Fields(1).Value)
                    If IsDBNull(chkrs1.Fields(2).Value) Then
                        TAD1 = ""
                    Else
                        TAD1 = chkrs1.Fields(2).Value
                    End If
                    If IsDBNull(chkrs1.Fields(3).Value) Then
                        TAD2 = ""
                    Else

                        TAD2 = chkrs1.Fields(3).Value
                    End If
                    If IsDBNull(chkrs1.Fields(4).Value) Then
                        TAD3 = ""
                    Else
                        TAD3 = chkrs1.Fields(4).Value
                    End If
                    If IsDBNull(chkrs1.Fields(5).Value) Then
                        TCITY = ""
                    Else
                        TCITY = chkrs1.Fields(5).Value
                    End If
                    If IsDBNull(chkrs1.Fields(17).Value) Then
                        TSTATE = ""
                    Else
                        TSTATE = chkrs1.Fields(17).Value
                    End If
                    STATE_CODE = "24"
                    If IsDBNull(chkrs1.Fields(18).Value) Then
                        TEMAIL = ""
                    Else
                        TEMAIL = chkrs1.Fields(18).Value
                    End If
                    If IsDBNull(chkrs1.Fields(19).Value) Then
                        TGST = ""
                    Else
                        TGST = chkrs1.Fields(19).Value
                    End If
                End If
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, "BILLED TO :" & Space(35) & "[  ] Original for Recepient" & vbNewLine)
                Print(fnum, "           " & Space(35) & "[  ] Duplicate for Supplier" & vbNewLine)
                Print(fnum, GetStringToPrint(45, TENANT_NAME, "S") & " " & StrDup(27, "-") & vbNewLine)
                If TAD1 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD1, "S") & vbNewLine)
                End If
                If TAD2 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD2, "S") & vbNewLine)
                End If
                If TAD3 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD3, "S") & vbNewLine)
                End If
                If TCITY <> "" Then
                    Print(fnum, GetStringToPrint(50, TCITY, "S") & vbNewLine)
                End If
                '  If TSTATE <> "" Then
                Print(fnum, GetStringToPrint(30, "STATE :" & TSTATE, "S") & Space(15) & "GODOWN NO.  :" & GetStringToPrint(20, chkrs.Fields(0).Value & chkrs.Fields(3).Value, "S") & vbNewLine)
                '  End If
                ' If TGST <> "" Then
                Print(fnum, "GSTIN :" & GetStringToPrint(30, TGST, "S") & Space(8) & "INVOICE NO. :GO-" & GetStringToPrint(35, INVOICE_NO, "S") & vbNewLine)
                '  End If
                '  If TEMAIL <> "" Then
                Print(fnum, "EMAIL ID:" & GetStringToPrint(33, TEMAIL, "S") & Space(3) & "INVOICE DATE:" & GetStringToPrint(20, DateTimePicker1.Value.ToString("dd/MM/yyyy"), "S") & vbNewLine)
                '   End If
                chkrs1.Close()
                Print(fnum, StrDup(30, " ") & vbNewLine)
                Print(fnum, StrDup(30, " ") & "TAX INVOICE FOR SERVICES" & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)



                Print(fnum, GetStringToPrint(8, "HSN", "S") & GetStringToPrint(28, "HSN DESCRIPTION", "S") & GetStringToPrint(30, "DESCRIPTION OF SERVICES", "S") & GetStringToPrint(19, "AMOUNT", "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                If IsDBNull(chkrs.Fields(37).Value) Then
                    Print(fnum, GetStringToPrint(7, "", "S"))
                Else
                    Print(fnum, GetStringToPrint(7, chkrs.Fields(37).Value, "S"))
                End If
                Print(fnum, GetStringToPrint(28, " Rental Or Leasing Services ", "S"))
                Print(fnum, GetStringToPrint(41, " Rent for property from " & "1st " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S"))

                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and P_CODE ='" & TENANT_CODE & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                amt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()

                    amt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amt = amt + chkrs2.Fields(5).Value
                    End If
                End If

                Print(fnum, GetStringToPrint(9, Format(amt, "#####0.00"), "N") & vbNewLine)

                chkrs2.Close()
                Dim ENDDAY As String
                ENDDAY = DateTime.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month).ToString
                If ENDDAY = "31" Then
                    ENDDAY = "31st"
                Else
                    ENDDAY = ENDDAY + "th"
                End If
                Print(fnum, GetStringToPrint(7, "", "S"))
                Print(fnum, GetStringToPrint(28, " Involving Own Or Leased ", "S"))
                Print(fnum, GetStringToPrint(35, " to " & ENDDAY & " " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S"))
                If IsDBNull(chkrs.Fields(37).Value) Or chkrs.Fields(37).Value.Equals("997211") Then
                    Print(fnum, GetStringToPrint(29, " Residential Property ", "S"))
                Else
                    Print(fnum, GetStringToPrint(29, " Non-residential Property ", "S"))
                End If
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "TAXABLE AMOUNT :", "S") & GetStringToPrint(10, Format(amt, "#####0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)

                chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs.Fields(37).Value & "'", xcon)
                CGST_RATE = 0
                SGST_RATE = 0
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
                gst_amt = gst * amt / 100



                Dim net As Double
                Dim rnd As Integer
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If

                net = amt + gst_amt
                CGST_TAXAMT = gst_amt / 2


                'CGST_TAXAMT = amt * CGST_RATE / 100
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                Print(fnum, Space(58) & GetStringToPrint(17, "CGST@ " & CGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(CGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "SGST@ " & SGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(SGST_TAXAMT, "######0.00"), "N") & vbNewLine)

                Print(fnum, Space(58) & GetStringToPrint(17, "NET AMOUNT     :", "S") & GetStringToPrint(10, Format(net, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(net)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(49, inwordd.Length - 49)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(23) & GetStringToPrint(50, inword1, "S") & vbNewLine)
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If

                Print(fnum, StrDup(85, "-") & vbNewLine)
                ''''''''''''''''''''''''''''''''SEARCH FOR ADVANCE START
                Dim adjamt As Double
                Dim advrec As Integer = 0
                Dim adv_date As Date
                chkrs2.Open("SELECT * FROM RECEIPT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and ADVANCE = TRUE AND ADJ_AMT>0 order by REC_DATE", xcon)
                adjamt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()
                    advrec = chkrs2.Fields(4).Value
                    DateTimePicker2.Value = chkrs2.Fields(3).Value
                    adv_date = chkrs2.Fields(3).Value
                    adjamt = chkrs2.Fields(13).Value - net
                    Dim sav As String = "UPDATE [RECEIPT] SET ADJ_AMT=" & adjamt & " WHERE REC_NO=" & chkrs2.Fields(4).Value & " AND year(REC_DATE)='" & Convert.ToDateTime(chkrs2.Fields(3).Value).Year & "'"
                    doSQL(sav)
                End If
                chkrs2.Close()
                ''''''''''''''''''''''''''''''''SEARCH FOR ADVANCE END
                Print(fnum, Space(40) & GetStringToPrint(45, "For Motilal Hirabhai Estate & Warehouse Ltd.", "S") & vbNewLine)
                If advrec > 0 Then
                    Print(fnum, GetStringToPrint(23, " ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(24, "Received as advance on:", "S") & GetStringToPrint(19, adv_date, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Receipt No.:", "S") & GetStringToPrint(19, "GST-" + Convert.ToString(advrec), "S") & vbNewLine)
                Else
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                End If

                Print(fnum, Space(40) & GetStringToPrint(45, "Authorised Signatory", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Subject to Ahmedabad jurisdiction.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "This is computer generated invoice.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Ref. No. :" + FILE_NOtmp, "S") & vbNewLine)
                Dim save As String
                If advrec > 0 Then
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO,REC_NO,REC_DATE,ADVANCE) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "','" & advrec & "','" & DateTimePicker2.Value.ToShortDateString & "',TRUE)"
                Else
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "')"
                End If
                doSQL(save)
                FileClose(fnum)
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
                End If
                strReportFilePath = Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat"
                CreatePDF(strReportFilePath, FILE_NO)
                If chkrs.EOF = False Then
                    chkrs.MoveNext()
                End If
                If chkrs.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs.Close()
            xcon.Close()

            'RichTextBox2.LoadFile(strReportFilePath, RichTextBoxStreamType.PlainText)
            MsgBox(numRec.ToString + " Bills are generated at " + Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) + "\ path")
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Function

    Private Function checkprlogo()
        Try
            Dim objPRNSetup = New clsPrinterSetup
            'set Paper Lines and Left Margin
            prnmaxpagelines = objPRNSetup.LinesPerPage
            If objPRNSetup.PageSize = PRNA4Paper Then
                prnleftmargin = 7
            Else
                prnleftmargin = 2
            End If
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If

            Dim array() As String = {"AE", "AA", "AB", "AC"}
            Dim subsql As String
            subsql = "SELECT * FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "'"
            chkrs1.Open("SELECT * FROM BILL where Month([BILL].bill_date)='" & Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' and Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "'", xcon)

            If chkrs1.EOF = False Then
                MsgBox("Bills for date " + DateTimePicker1.Value.ToString + " are already generated ")
                chkrs1.Close()
                Exit Function
            End If
            chkrs1.Close()
            Dim iDate As String
            Dim fDate As DateTime
            Dim oDate As String
            Dim foDate As DateTime
            If (Month(Convert.ToDateTime(DateTimePicker1.Value.ToString)) >= 4) Then
                iDate = "30/04/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) + 1))
                foDate = Convert.ToDateTime(oDate)
            Else
                iDate = "30/04/" + Convert.ToString((Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) - 1))
                fDate = Convert.ToDateTime(iDate)
                oDate = "31/03/" + Convert.ToString(Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)))
                foDate = Convert.ToDateTime(oDate)
            End If

            ' MsgBox(fDate.ToString)
            ' MsgBox(foDate.ToString)
            Dim strn As String = "Select * FROM BILL where [BILL].bill_date>=Format('" & fDate & "', 'dd/mm/yyyy') and [BILL].bill_date<=Format('" & foDate & "', 'dd/mm/yyyy') order by Year([BILL].bill_date)+[BILL].INVOICE_NO"
            '''''''''''''get last serial number
            '''''chkrs11.Open("Select * FROM BILL where Year([BILL].bill_date)='" & Year(Convert.ToDateTime(DateTimePicker1.Value.ToString)) & "' order by Year([BILL].bill_date)+[BILL].INVOICE_NO", xcon)
            chkrs11.Open(strn, xcon)
            Dim srno As Integer = 0
            Do While chkrs11.EOF = False
                srno = chkrs11.Fields(0).Value
                If chkrs11.EOF = False Then
                    chkrs11.MoveNext()
                End If
            Loop
            chkrs11.Close()


            '''''''''''''get last serial number

            fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
            xcount = 0      '''''''''Set xcount'''''''''''''''''
            xlimit = 88     '''''''''Set xlimit'''''''''''''''''

            'xpage = 1
            xpage = Val("2")
            Dim i1 As Double
            ''''''''''''''''' open a file sharereg.txt'''''''''''
            ' FileOpen(fnum, Application.StartupPath & "\Invoices\RecordSlipView.dat", OpenMode.Output)
            '  Call header()
            Dim numRec As Integer = 0
            Dim STRR As String = "SELECT * FROM GODOWN where [STATUS]='C' AND FROM_D <=#" & Convert.ToDateTime(DateTimePicker1.Value) & "# order by [GROUP]+GODWN_NO "
            chkrs.Open("SELECT * FROM GODOWN where [STATUS]='C' AND FROM_D <=FORMAT('" & Convert.ToDateTime(DateTimePicker1.Value) & "','DD/MM/YYYY') order by [GROUP]+GODWN_NO ", xcon)
            Do While chkrs.EOF = False
                numRec = numRec + 1
                srno = srno + 1
                Dim INVOICE_NO As String
                Dim FILE_NO As String
                Dim FILE_NOtmp As String
                'INVOICE_NO = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & chkrs.Fields(0).Value & chkrs.Fields(3).Value     'String.Format("{0:000}", srno)
                FILE_NOtmp = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", numRec)
                ' FILE_NOtmp = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & chkrs.Fields(0).Value & chkrs.Fields(3).Value     'String.Format("{0:000}", srno)
                'INVOICE_NO = DateTimePicker1.Value.Year & "-" & Convert.ToInt32(DateTimePicker1.Value.Year.ToString.Substring(2)) + 1 & "/" & MonthName(DateTimePicker1.Value.Month, True) & "/" & String.Format("{0:000}", srno)
                INVOICE_NO = String.Format("{0:0000}", srno)
                FILE_NO = INVOICE_NO.Replace("/", "_")   'DateTimePicker1.Value.Year & "_" & DateTimePicker1.Value.Month & "_" & chkrs.Fields(0).Value & chkrs.Fields(3).Value.replace("/", "_")
                FILE_NO = FILE_NO.Replace(" ", "_")
                ' FILE_NO = FILE_NO.Replace("-", "_")
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
                End If


                FileOpen(fnum, Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat", OpenMode.Output)
                Dim TENANT_CODE As String
                Dim TENANT_NAME As String
                Dim T_ADDREASS As String
                Dim TAD1 As String
                Dim TAD2 As String
                Dim TAD3 As String
                Dim TCITY As String
                Dim TSTATE As String
                Dim STATE_CODE As String
                Dim TEMAIL As String
                Dim TGST As String
                Dim amt As Double
                Dim CGST_TAXAMT As Double
                Dim SGST_TAXAMT As Double
                Dim CGST_RATE As Double
                Dim gst_amt As Double
                Dim SGST_RATE As Double
                Dim gst As Double
                TENANT_CODE = chkrs.Fields(1).Value
                chkrs1.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
                If chkrs1.EOF = False Then
                    TENANT_NAME = LTrim(chkrs1.Fields(1).Value)
                    If IsDBNull(chkrs1.Fields(2).Value) Then
                        TAD1 = ""
                    Else
                        TAD1 = chkrs1.Fields(2).Value
                    End If
                    If IsDBNull(chkrs1.Fields(3).Value) Then
                        TAD2 = ""
                    Else

                        TAD2 = chkrs1.Fields(3).Value
                    End If
                    If IsDBNull(chkrs1.Fields(4).Value) Then
                        TAD3 = ""
                    Else
                        TAD3 = chkrs1.Fields(4).Value
                    End If
                    If IsDBNull(chkrs1.Fields(5).Value) Then
                        TCITY = ""
                    Else
                        TCITY = chkrs1.Fields(5).Value
                    End If
                    If IsDBNull(chkrs1.Fields(17).Value) Then
                        TSTATE = ""
                    Else
                        TSTATE = chkrs1.Fields(17).Value
                    End If
                    STATE_CODE = "24"
                    If IsDBNull(chkrs1.Fields(18).Value) Then
                        TEMAIL = ""
                    Else
                        TEMAIL = chkrs1.Fields(18).Value
                    End If
                    If IsDBNull(chkrs1.Fields(19).Value) Then
                        TGST = ""
                    Else
                        TGST = chkrs1.Fields(19).Value
                    End If
                End If
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, StrDup(28, " ") & vbNewLine)
                Print(fnum, "BILLED TO :" & Space(35) & "[  ] Original for Recepient" & vbNewLine)
                Print(fnum, "           " & Space(35) & "[  ] Duplicate for Supplier" & vbNewLine)
                Print(fnum, GetStringToPrint(45, TENANT_NAME, "S") & " " & StrDup(27, "-") & vbNewLine)
                If TAD1 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD1, "S") & vbNewLine)
                End If
                If TAD2 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD2, "S") & vbNewLine)
                End If
                If TAD3 <> "" Then
                    Print(fnum, GetStringToPrint(50, TAD3, "S") & vbNewLine)
                End If
                If TCITY <> "" Then
                    Print(fnum, GetStringToPrint(50, TCITY, "S") & vbNewLine)
                End If
                '  If TSTATE <> "" Then
                Print(fnum, GetStringToPrint(30, "STATE :" & TSTATE, "S") & Space(15) & "GODOWN NO.  :" & GetStringToPrint(20, chkrs.Fields(0).Value & chkrs.Fields(3).Value, "S") & vbNewLine)
                '  End If
                ' If TGST <> "" Then
                Print(fnum, "GSTIN :" & GetStringToPrint(30, TGST, "S") & Space(8) & "INVOICE NO. :GO-" & GetStringToPrint(35, INVOICE_NO, "S") & vbNewLine)
                '  End If
                '  If TEMAIL <> "" Then
                Print(fnum, "EMAIL ID:" & GetStringToPrint(33, TEMAIL, "S") & Space(3) & "INVOICE DATE:" & GetStringToPrint(20, DateTimePicker1.Value.ToString("dd/MM/yyyy"), "S") & vbNewLine)
                '   End If
                chkrs1.Close()
                Print(fnum, StrDup(30, " ") & vbNewLine)
                Print(fnum, StrDup(30, " ") & "TAX INVOICE FOR SERVICES" & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)



                Print(fnum, GetStringToPrint(8, "HSN", "S") & GetStringToPrint(28, "HSN DESCRIPTION", "S") & GetStringToPrint(30, "DESCRIPTION OF SERVICES", "S") & GetStringToPrint(19, "AMOUNT", "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                If IsDBNull(chkrs.Fields(37).Value) Then
                    Print(fnum, GetStringToPrint(7, "", "S"))
                Else
                    Print(fnum, GetStringToPrint(7, chkrs.Fields(37).Value, "S"))
                End If
                Print(fnum, GetStringToPrint(28, " Rental Or Leasing Services ", "S"))
                Print(fnum, GetStringToPrint(41, " Rent for property from " & "1st " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S"))

                chkrs2.Open("SELECT * FROM RENT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and P_CODE ='" & TENANT_CODE & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
                amt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()

                    amt = chkrs2.Fields(4).Value
                    If IsDBNull(chkrs2.Fields(5).Value) Then
                    Else
                        amt = amt + chkrs2.Fields(5).Value
                    End If
                End If

                Print(fnum, GetStringToPrint(9, Format(amt, "#####0.00"), "N") & vbNewLine)

                chkrs2.Close()
                Dim ENDDAY As String
                ENDDAY = DateTime.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month).ToString
                If ENDDAY = "31" Then
                    ENDDAY = "31st"
                Else
                    ENDDAY = ENDDAY + "th"
                End If
                Print(fnum, GetStringToPrint(7, "", "S"))
                Print(fnum, GetStringToPrint(28, " Involving Own Or Leased ", "S"))
                Print(fnum, GetStringToPrint(35, " to " & ENDDAY & " " & MonthName(DateTimePicker1.Value.Month) & "," & DateTimePicker1.Value.Year.ToString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S"))
                If IsDBNull(chkrs.Fields(37).Value) Or chkrs.Fields(37).Value.Equals("997211") Then
                    Print(fnum, GetStringToPrint(29, " Residential Property ", "S"))
                Else
                    Print(fnum, GetStringToPrint(29, " Non-residential Property ", "S"))
                End If
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "TAXABLE AMOUNT :", "S") & GetStringToPrint(10, Format(amt, "#####0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)

                chkrs3.Open("SELECT * FROM GST WHERE [HSN_NO]='" & chkrs.Fields(37).Value & "'", xcon)
                CGST_RATE = 0
                SGST_RATE = 0
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
                gst_amt = gst * amt / 100



                Dim net As Double
                Dim rnd As Integer
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If

                net = amt + gst_amt
                CGST_TAXAMT = gst_amt / 2


                'CGST_TAXAMT = amt * CGST_RATE / 100
                CGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                SGST_TAXAMT = Math.Round(gst_amt / 2, 2, MidpointRounding.AwayFromZero)
                Print(fnum, Space(58) & GetStringToPrint(17, "CGST@ " & CGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(CGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "SGST@ " & SGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(SGST_TAXAMT, "######0.00"), "N") & vbNewLine)

                Print(fnum, Space(58) & GetStringToPrint(17, "NET AMOUNT     :", "S") & GetStringToPrint(10, Format(net, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(net)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(49, inwordd.Length - 49)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(23) & GetStringToPrint(50, inword1, "S") & vbNewLine)
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If

                Print(fnum, StrDup(85, "-") & vbNewLine)
                ''''''''''''''''''''''''''''''''SEARCH FOR ADVANCE START
                Dim adjamt As Double
                Dim advrec As Integer = 0
                Dim adv_date As Date
                chkrs2.Open("SELECT * FROM RECEIPT WHERE [GROUP]='" & chkrs.Fields(0).Value & "' and GODWN_NO='" & chkrs.Fields(3).Value & "' and ADVANCE = TRUE AND ADJ_AMT>0 order by REC_DATE", xcon)
                adjamt = 0
                If chkrs2.EOF = False Then
                    chkrs2.MoveFirst()
                    advrec = chkrs2.Fields(4).Value
                    DateTimePicker2.Value = chkrs2.Fields(3).Value
                    adv_date = chkrs2.Fields(3).Value
                    adjamt = chkrs2.Fields(13).Value - net
                    Dim sav As String = "UPDATE [RECEIPT] SET ADJ_AMT=" & adjamt & " WHERE REC_NO=" & chkrs2.Fields(4).Value & " AND year(REC_DATE)='" & Convert.ToDateTime(chkrs2.Fields(3).Value).Year & "'"
                    doSQL(sav)
                End If
                chkrs2.Close()
                ''''''''''''''''''''''''''''''''SEARCH FOR ADVANCE END
                Print(fnum, Space(40) & GetStringToPrint(45, "For Motilal Hirabhai Estate & Warehouse Ltd.", "S") & vbNewLine)
                If advrec > 0 Then
                    Print(fnum, GetStringToPrint(23, " ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(24, "Received as advance on:", "S") & GetStringToPrint(19, adv_date, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Receipt No.:", "S") & GetStringToPrint(19, "GST-" + Convert.ToString(advrec), "S") & vbNewLine)
                Else
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                End If

                Print(fnum, Space(40) & GetStringToPrint(45, "Authorised Signatory", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Subject to Ahmedabad jurisdiction.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "This is computer generated invoice.", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, " ", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(80, "Ref. No. :" + FILE_NOtmp, "S") & vbNewLine)
                Dim save As String
                If advrec > 0 Then
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO,REC_NO,REC_DATE,ADVANCE) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "','" & advrec & "','" & DateTimePicker2.Value.ToShortDateString & "',TRUE)"
                Else
                    save = "INSERT INTO [BILL](INVOICE_NO,[GROUP],GODWN_NO,P_CODE,BILL_DATE,BILL_AMOUNT,CGST_RATE,CGST_AMT,SGST_RATE,SGST_AMT,NET_AMOUNT,HSN,SRNO) VALUES('" & INVOICE_NO & "','" & chkrs.Fields(0).Value & "','" & chkrs.Fields(3).Value & "','" & chkrs.Fields(1).Value & "','" & Convert.ToDateTime(DateTimePicker1.Value.ToString) & "'," & amt & "," & CGST_RATE & "," & CGST_TAXAMT & "," & SGST_RATE & "," & SGST_TAXAMT & "," & net & ",'" & chkrs.Fields(37).Value & "','" & FILE_NOtmp & "')"
                End If
                doSQL(save)
                FileClose(fnum)
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month))
                End If
                strReportFilePath = Application.StartupPath & "\Invoices\dat\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & FILE_NO & ".dat"
                CreatePDF(strReportFilePath, FILE_NO)
                If chkrs.EOF = False Then
                    chkrs.MoveNext()
                End If
                If chkrs.EOF = True Then
                    Exit Do
                End If
            Loop
            chkrs.Close()
            xcon.Close()

            'RichTextBox2.LoadFile(strReportFilePath, RichTextBoxStreamType.PlainText)
            MsgBox(numRec.ToString + " Bills are generated at " + Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) + "\ path")
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Function

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
            ' MsgBox("Data Inserted successfully in database", vbInformation)
            objcmd.Dispose()
            MyConn.Close()
        Catch ex As Exception
            MsgBox("Exception: Data Insertion in PARTY table in database" & ex.Message)
        End Try
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 0

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"
            Dim pdfPage As PdfPage = pdf.AddPage
            pdfPage.TrimMargins.Left = 15

            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)


            Dim font As XFont = New XFont("COURIER NEW", 9, XFontStyle.Regular)
            If ChkLogo.Checked Then
                Dim image As XImage = image.FromFile(Application.StartupPath & "\logo.png")
                graph.DrawImage(image, 0, 0, image.Width, image.Height)
            End If
            While True
                line = readFile.ReadLine()
                If line Is Nothing Then
                    Exit While
                Else

                    ' yPoint = yPoint + 1
                    graph.DrawString(line, font, XBrushes.Black,
                     New XRect(50, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = Application.StartupPath & "\Invoices\pdf\" & DateTimePicker1.Value.Year.ToString & "\" & MonthName(DateTimePicker1.Value.Month) & "\" & invoice_no & ".pdf"

            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With Me
            .Cursor = Cursors.WaitCursor
            .Refresh()
        End With
        checkpr()
        With Me
            .Cursor = Cursors.Default
            .Refresh()
        End With
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim selectedDate As Date = DateTimePicker1.Value
        Dim DaysInMonth As Integer = Date.DaysInMonth(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month)
        If (DateTimePicker1.Value.Day <> DaysInMonth) Then
            DateTimePicker1.Value = Convert.ToDateTime(DaysInMonth.ToString + "/" + DateTimePicker1.Value.Month.ToString + "/" + DateTimePicker1.Value.Year.ToString)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmInvoice_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Dim MyDate As Date = Now
            Dim DaysInMonth As Integer
            If DateTime.Now.Month = 1 Then
                DaysInMonth = Date.DaysInMonth(MyDate.Year - 1, 12)
                DateTimePicker1.Value = Convert.ToDateTime(DaysInMonth.ToString + "/" + (12).ToString + "/" + (MyDate.Year - 1).ToString)
            Else
                DaysInMonth = Date.DaysInMonth(MyDate.Year, MyDate.Month - 1)
                DateTimePicker1.Value = Convert.ToDateTime(DaysInMonth.ToString + "/" + (MyDate.Month - 1).ToString + "/" + MyDate.Year.ToString)
            End If

            'If (DateTimePicker1.Value.Day <> DaysInMonth) Then
            '    DateTimePicker1.Value = Convert.ToDateTime(DaysInMonth.ToString + "/" + (MyDate.Month - 1).ToString + "/" + MyDate.Year.ToString)
            'End If
            formloaded = True
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmInvoice_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub ChkLogo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkLogo.CheckedChanged

    End Sub
End Class