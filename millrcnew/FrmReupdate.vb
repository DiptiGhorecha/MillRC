Option Explicit On
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmReupdate
    Dim chkrs As New ADODB.Recordset
    Dim chkrs11 As New ADODB.Recordset
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim chkrs3 As New ADODB.Recordset
    Dim chkrs4 As New ADODB.Recordset
    Dim chkrs5 As New ADODB.Recordset
    Dim chkrs6 As New ADODB.Recordset
    Dim chkrs22 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim xrs As New ADODB.Recordset      '''''''' variable is use to open a Recordset
    Dim xtemp As New ADODB.Recordset    '''''''' used to open a temparory Recordset
    Dim XComp As New ADODB.Recordset
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;"
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim fnum As Integer
    Dim fnumm As Integer
    Dim pdfpath As String
    Dim strReportFilePath As String
    Private Sub FrmReupdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Dim MyDate As Date = Now
            Dim DaysInMonth As Integer
            formloaded = True
        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        With Me
            .Cursor = Cursors.WaitCursor
            .Refresh()
        End With

        Try
            MyConn = New OleDbConnection(connString)
            If MyConn.State = ConnectionState.Closed Then
                MyConn.Open()
            End If
            Dim transaction As OleDbTransaction
            transaction = MyConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Dim objcmd As New OleDb.OleDbCommand
            Dim objcmdd As New OleDb.OleDbCommand

            objcmd.Connection = MyConn
            objcmd.Transaction = transaction
            objcmd.CommandType = CommandType.Text
            Dim save As String = "UPDATE [BILL] SET REC_NO=Null, REC_DATE=Null ,ADVANCE=FALSE WHERE INVOICE_NO<>'" & "" & "'"
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()
            objcmdd.Connection = MyConn
            objcmdd.Transaction = transaction
            objcmdd.CommandType = CommandType.Text
            save = "UPDATE [RECEIPT] SET ADJ_AMT=AMOUNT WHERE REC_NO<>0"
            objcmd.CommandText = save
            objcmd.ExecuteNonQuery()
            objcmd.CommandText = "delete * from [RECIPTBILL] where REC_NO<>0"
            objcmd.ExecuteNonQuery()
            transaction.Commit()
            MyConn.Close()
            If xcon.State = ConnectionState.Open Then
            Else
                xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
            End If
            Dim AMT As Double = 0
            Dim ADJUSTED As Double = 0
            Dim RCDT As Date
            Dim RCNO As Integer
            Dim REMAINING As Double = 0
            Dim BLDT As Date
            Dim GRP As String
            Dim GDN As String
            Dim PCD As String
            Dim INVNO As String
            Dim hsn As String
            Dim BLAMT As Double
            Dim ADV As Boolean
            Dim gamt As Double
            Dim CGST_AMT As Double
            Dim SGST_AMT As Double
            Dim CGST_RT As Double
            Dim SGST_RT As Double
            Dim FILE_NOtmp As String
            Dim fdate As Date
            Dim radate As Date
            Dim rano As Integer
            chkrs2.Open("Select Count(invoice_no) As NumberOfInvoice FROM BILL", xcon)
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = chkrs2.Fields(0).Value
            chkrs2.Close()
            ProgressBar1.Value = 0
            Dim counter As Integer = 0
            chkrs1.Open("SELECT * FROM BILL ORDER BY BILL_DATE,INVOICE_NO", xcon)
            Do While chkrs1.EOF = False
                BLDT = chkrs1.Fields(4).Value
                GRP = chkrs1.Fields(1).Value
                GDN = chkrs1.Fields(2).Value
                PCD = chkrs1.Fields(3).Value
                INVNO = chkrs1.Fields(0).Value
                BLAMT = chkrs1.Fields(10).Value
                hsn = chkrs1.Fields(11).Value
                gamt = chkrs1.Fields(5).Value
                CGST_AMT = chkrs1.Fields(7).Value
                SGST_AMT = chkrs1.Fields(9).Value
                CGST_RT = chkrs1.Fields(6).Value
                SGST_RT = chkrs1.Fields(8).Value
                FILE_NOtmp = chkrs1.Fields(12).Value
                fdate = Nothing
                radate = Nothing
                rano = 0
                ADV = False
                ''''''''''''''''''''''''''opening advance
                chkrs22.Open("SELECT [ADVANCES].* from [ADVANCES] WHERE [GROUP]='" & GRP & "' AND [GODWN_NO]='" & GDN & "' AND [ADVANCES].P_CODE='" & PCD & "' order by [advances].[GROUP],[advances].GODWN_NO", xcon)
                While Not chkrs22.EOF
                    If (chkrs22.Fields(0).Value = GRP And chkrs22.Fields(1).Value = GDN) Then
                        ADV = True
                        fdate = chkrs22.Fields(3).Value
                        radate = chkrs22.Fields(4).Value
                        rano = chkrs22.Fields(5).Value
                    End If
                    If chkrs22.EOF = False Then
                        chkrs22.MoveNext()
                    End If
                    If chkrs22.EOF = True Then
                        Exit While
                    End If
                End While
                chkrs22.Close()
                Dim tstdt As Date = Nothing
                If (fdate <> Nothing) Then
                    tstdt = fdate
                End If
                '''''''''''''''''''''''''opening 
                If (BLDT <= tstdt) Then
                    MyConn = New OleDbConnection(connString)
                    If MyConn.State = ConnectionState.Closed Then
                        MyConn.Open()
                    End If
                    transaction = MyConn.BeginTransaction(IsolationLevel.ReadUncommitted)
                    Dim objcmd1 As New OleDb.OleDbCommand
                    objcmd1.Connection = MyConn
                    objcmd1.Transaction = transaction
                    objcmd1.CommandType = CommandType.Text
                    save = "UPDATE [BILL] SET REC_NO=" & rano & ", REC_DATE=format('" & Convert.ToDateTime(radate) & "','dd/mm/yyyy'),ADVANCE=" & ADV & " WHERE INVOICE_NO='" & INVNO & "' AND [GROUP]='" & GRP & "' AND GODWN_NO='" & GDN & "' AND BILL_DATE=format('" & Convert.ToDateTime(BLDT) & "','dd/mm/yyyy')"  ''' sorry about that
                    objcmd1.CommandText = save
                    objcmd1.ExecuteNonQuery()
                    transaction.Commit()
                    objcmd1.Dispose()
                    MyConn.Close()
                Else
                    Dim str As String = "SELECT * FROM RECEIPT WHERE GROUP='" & GRP & "' AND GODWN_NO='" & GDN & "' AND ADJ_AMT>0 ORDER BY [GROUP],GODWN_NO,REC_DATE,REC_NO"
                    chkrs2.Open("SELECT TOP 1 * FROM RECEIPT WHERE [GROUP]='" & GRP & "' AND GODWN_NO='" & GDN & "' AND ADJ_AMT>0 ORDER BY [GROUP],GODWN_NO,REC_DATE,REC_NO", xcon)
                    If chkrs2.EOF = False Then
                        ' chkrs2.MoveFirst()
                        RCDT = chkrs2.Fields(3).Value
                        RCNO = chkrs2.Fields(4).Value
                        REMAINING = chkrs2.Fields(13).Value - BLAMT
                        If (RCDT < BLDT) Then
                            ADV = True
                        Else
                            ADV = False
                        End If
                        MyConn = New OleDbConnection(connString)
                        If MyConn.State = ConnectionState.Closed Then
                            MyConn.Open()
                        End If
                        transaction = MyConn.BeginTransaction(IsolationLevel.ReadUncommitted)
                        Dim objcmd1 As New OleDb.OleDbCommand
                        objcmd1.Connection = MyConn
                        objcmd1.Transaction = transaction
                        objcmd1.CommandType = CommandType.Text
                        If ADV = False Then
                            objcmd1.CommandText = "INSERT INTO [RECIPTBILL](REC_NO,INVOICE_NO,AMT,REC_DATE) VALUES('" & RCNO & "','" & INVNO & "','" & BLAMT & "',Format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy'))"
                            objcmd1.ExecuteNonQuery()
                        End If
                        save = "UPDATE [BILL] SET REC_NO=" & RCNO & ", REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy'),ADVANCE=" & ADV & " WHERE INVOICE_NO='" & INVNO & "' AND [GROUP]='" & GRP & "' AND GODWN_NO='" & GDN & "' AND BILL_DATE=format('" & Convert.ToDateTime(BLDT) & "','dd/mm/yyyy')"  ''' sorry about that
                        objcmd1.CommandText = save
                        objcmd1.ExecuteNonQuery()
                        ' str = "UPDATE [RECEIPT] SET ADJ_AMT=" & REMAINING & " WHERE REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy') AND REC_NO=" & RCNO
                        objcmd1.CommandText = "UPDATE [RECEIPT] SET ADJ_AMT=" & REMAINING & " WHERE REC_DATE=format('" & Convert.ToDateTime(RCDT) & "','dd/mm/yyyy') AND REC_NO=" & RCNO  ' sorry about that
                        objcmd1.ExecuteNonQuery()
                        transaction.Commit()
                        objcmd1.Dispose()
                        MyConn.Close()
                        '''''' System.Threading.Thread.Sleep(500)
                    End If
                    chkrs2.Close()
                End If
                'Dim i As Integer
                'For i = 1 To 5000

                'Next

                ''''''''''''''''''''''''''''''''invoice reprint start
                fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
                If (Not System.IO.Directory.Exists(Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month))) Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month))
                End If
                FileOpen(fnum, Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month) & "\" & INVNO & ".dat", OpenMode.Output)
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
                Dim CGST_TAXAMT As Double
                Dim SGST_TAXAMT As Double
                Dim CGST_RATE As Double
                Dim SGST_RATE As Double
                Dim gst As Double
                Dim gst_amt As Double
                TENANT_CODE = PCD

                chkrs11.Open("SELECT * FROM PARTY WHERE P_CODE='" & TENANT_CODE & "'", xcon)
                If chkrs11.EOF = False Then
                    TENANT_NAME = LTrim(chkrs11.Fields(1).Value)

                    If IsDBNull(chkrs11.Fields(2).Value) Then
                        TAD1 = ""
                    Else
                        If Trim(chkrs11.Fields(2).Value).Equals("") Then
                            TAD1 = ""
                        Else
                            TAD1 = chkrs11.Fields(2).Value.replace("& vbLf & vbLf", "")
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(3).Value) Then
                        TAD2 = ""
                    Else
                        If Trim(chkrs11.Fields(3).Value).Equals("") Then
                            TAD2 = ""
                        Else
                            TAD2 = chkrs11.Fields(3).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(4).Value) Then
                        TAD3 = ""
                    Else
                        If Trim(chkrs11.Fields(4).Value).Equals("") Then
                            TAD3 = ""
                        Else
                            TAD3 = chkrs11.Fields(4).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(5).Value) Then
                        TCITY = ""
                    Else
                        If Trim(chkrs11.Fields(5).Value).Equals("") Then
                            TCITY = ""
                        Else
                            TCITY = chkrs11.Fields(5).Value
                        End If
                    End If
                    If IsDBNull(chkrs11.Fields(17).Value) Then
                        TSTATE = ""
                    Else
                        If Trim(chkrs11.Fields(17).Value).Equals("") Then
                            TSTATE = ""
                        Else
                            TSTATE = chkrs11.Fields(17).Value
                        End If
                    End If
                    STATE_CODE = "24"

                    If IsDBNull(chkrs11.Fields(18).Value) Then
                        TEMAIL = ""
                    Else
                        TEMAIL = chkrs11.Fields(18).Value
                    End If
                    If IsDBNull(chkrs11.Fields(19).Value) Then
                        TGST = ""
                    Else
                        TGST = chkrs11.Fields(19).Value
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
                Print(fnum, GetStringToPrint(30, "STATE :" & TSTATE, "S") & Space(15) & "GODOWN NO.  :" & GetStringToPrint(20, GRP & GDN, "S") & vbNewLine)
                '  End If
                ' If TGST <> "" Then
                Print(fnum, "GSTIN :" & GetStringToPrint(30, TGST, "S") & Space(8) & "INVOICE NO. :" & GetStringToPrint(35, "GO-" & INVNO, "S") & vbNewLine)
                '  End If
                '  If TEMAIL <> "" Then
                Print(fnum, "EMAIL ID:" & GetStringToPrint(33, TEMAIL, "S") & Space(3) & "INVOICE DATE:" & GetStringToPrint(20, BLDT.ToString("dd/MM/yyyy"), "S") & vbNewLine)
                '   End If
                chkrs11.Close()

                Print(fnum, StrDup(30, " ") & vbNewLine)
                Print(fnum, StrDup(30, " ") & "TAX INVOICE FOR SERVICES" & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, GetStringToPrint(8, "HSN", "S") & GetStringToPrint(28, "HSN DESCRIPTION", "S") & GetStringToPrint(30, "DESCRIPTION OF SERVICES", "S") & GetStringToPrint(19, "AMOUNT", "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                If IsDBNull(hsn) Then
                    Print(fnum, GetStringToPrint(7, "", "S"))
                Else
                    Print(fnum, GetStringToPrint(7, hsn, "S"))

                End If
                Print(fnum, GetStringToPrint(28, " Rental Or Leasing Services ", "S"))
                Print(fnum, GetStringToPrint(41, " Rent for property from " & "1st " & MonthName(BLDT.Month) & "," & BLDT.Year.ToString, "S"))
                '  ComboBox1.Text = DateTimePicker1.Value.Month
                '   ComboBox2.Text = DateTimePicker1.Value.Year

                ' AMT = TextBox4.Text
                Print(fnum, GetStringToPrint(9, Format(gamt, "#####0.00"), "N") & vbNewLine)
                Dim ENDDAY As String
                ENDDAY = DateTime.DaysInMonth(BLDT.Year, BLDT.Month).ToString
                If ENDDAY = "31" Then
                    ENDDAY = "31st"
                Else
                    ENDDAY = ENDDAY + "th"
                End If
                Print(fnum, GetStringToPrint(7, "", "S"))
                Print(fnum, GetStringToPrint(28, " Involving Own Or Leased ", "S"))
                Print(fnum, GetStringToPrint(35, " to " & ENDDAY & " " & MonthName(BLDT.Month) & "," & BLDT.Year.ToString, "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S"))
                If IsDBNull(hsn) Or hsn.Equals("997211") Then
                    Print(fnum, GetStringToPrint(29, " Residential Property ", "S"))
                    '    Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Residential Property"
                Else
                    Print(fnum, GetStringToPrint(29, " Non-residential Property ", "S"))
                    '   Label18.Text = "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"
                End If

                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, GetStringToPrint(7, "", "S") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "TAXABLE AMOUNT :", "S") & GetStringToPrint(10, Format(gamt, "#####0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)

                'CGST_TAXAMT = amt * CGST_RATE / 100
                'CGST_TAXAMT = Math.Round(CGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
                'SGST_TAXAMT = amt * SGST_RATE / 100
                'SGST_TAXAMT = Math.Round(SGST_TAXAMT, 2, MidpointRounding.AwayFromZero)
                Dim net As Double
                Dim rnd As Integer


                net = BLAMT
                CGST_TAXAMT = CGST_AMT
                SGST_TAXAMT = SGST_AMT
                CGST_RATE = CGST_RT
                SGST_RATE = SGST_RT

                Print(fnum, Space(58) & GetStringToPrint(17, "CGST@ " & CGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(CGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, Space(58) & GetStringToPrint(17, "SGST@ " & SGST_RATE & "%" & "       :", "S") & GetStringToPrint(10, Format(SGST_TAXAMT, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(58, " ") & StrDup(27, "-") & vbNewLine)

                Print(fnum, Space(58) & GetStringToPrint(17, "NET AMOUNT     :", "S") & GetStringToPrint(10, Format(net, "######0.00"), "N") & vbNewLine)
                Print(fnum, StrDup(85, "-") & vbNewLine)
                Dim inwordd As String = ""
                Dim inword As String = ""
                Dim inword1 As String = ""
                inwordd = convinRS(net)
                If inwordd.Length > 50 Then
                    inword = inwordd.Substring(0, 49)
                    inword1 = inwordd.Substring(50, inwordd.Length - 50)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    Print(fnum, Space(23) & GetStringToPrint(51, inword1, "S") & vbNewLine)
                Else
                    inword = inwordd.Substring(0, inwordd.Length)
                    Print(fnum, GetStringToPrint(35, "Amount Chargeable (In Words): INR ", "S") & GetStringToPrint(50, inword, "S") & vbNewLine)
                    'Print(fnum, Space(23) & GetStringToPrint(61, inword1, "S") & vbNewLine)
                End If

                Print(fnum, StrDup(85, "-") & vbNewLine)
                Print(fnum, Space(40) & GetStringToPrint(45, "For Motilal Hirabhai Estate & Warehouse Ltd.", "S") & vbNewLine)
                If ADV = True Then
                    Print(fnum, GetStringToPrint(23, " ", "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(24, "Received as advance on:", "S") & GetStringToPrint(19, RCDT, "S") & vbNewLine)
                    Print(fnum, GetStringToPrint(13, "Receipt No.:", "S") & GetStringToPrint(19, "GST-" + Convert.ToString(RCNO), "S") & vbNewLine)
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
                FileClose(fnum)
                pdfpath = Application.StartupPath & "\Invoices\pdf\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month)
                strReportFilePath = Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month) & "\" & INVNO & ".dat"
                CreatePDFNEW(strReportFilePath, INVNO)
                ''''''''''''''''''''''''''''''''invoice reprint end
                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
                ProgressBar1.Value += 1
            Loop
            chkrs1.Close()
            xcon.Close()

        Catch ex As Exception
            MessageBox.Show("Error opening file-sr: " & ex.Message)
        End Try

        With Me
            .Cursor = Cursors.Default
            .Refresh()
        End With
        MsgBox("Reupdate process completed")
    End Sub
    Private Function CreatePDFNEW(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New System.IO.StreamReader(strReportFilePath)
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
                    graph.DrawString(line, font, XBrushes.Black,
                    New XRect(50, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = pdfpath & "\" & invoice_no & ".pdf"

            pdf.Save(pdfFilename)
            readFile.Close()
            readFile = Nothing
            ' Process.Start(pdfFilename)
            pdf.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

    Private Sub FrmReupdate_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
End Class