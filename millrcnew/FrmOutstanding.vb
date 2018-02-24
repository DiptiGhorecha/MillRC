
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Public Class FrmOutstanding
    Dim chkrs As New ADODB.Recordset
    Dim chkrs11 As New ADODB.Recordset
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
    Dim indexorder As String = "GODWN_NO"
    Dim ctrlname As String = "TextBox1"
    Dim formloaded As Boolean = False
    Dim fnum As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\outstand.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        If RadioGodown.Checked = True Then
            chkrs1.Open("SELECT [godown].*,[PARTY].P_NAME from [godown] INNER JOIN [PARTY] ON [godown].P_CODE=[PARTY].P_CODE WHERE [STATUS]='C' order by [godown].[GROUP],[godown].GODWN_NO,[godown].P_CODE", xcon)
        Else
            chkrs1.Open("SELECT [godown].*,[PARTY].P_NAME from [godown] INNER JOIN [PARTY] ON [godown].P_CODE=[PARTY].P_CODE WHERE [STATUS]='C' order by [PARTY].P_NAME", xcon)
        End If
        Dim gtotamt As Double = 0
        Dim gtotadv As Double = 0
        Dim firstrec As Boolean = True
        Do While chkrs1.EOF = False
            Dim grp As String = chkrs1.Fields(0).Value
            Dim gdn As String = chkrs1.Fields(3).Value
            Dim pcd As String = chkrs1.Fields(1).Value
            Dim pnm As String = chkrs1.Fields(38).Value
            Dim fdate As String = chkrs1.Fields(11).Value
            Dim pgst As String = chkrs1.Fields(37).Value
            Dim pgross As Double = 0
            Dim pcgst As Double = 0
            Dim psgst As Double = 0
            Dim pnet As Double = 0
            Dim rnet As Double = 0
            Dim out As Double = 0
            Dim lastBillGen As Date
            Dim monthamt As Double
            Dim iDate As String = "31/07/2017"
            Dim oDate As DateTime = Convert.ToDateTime(iDate)
            Dim outDate As DateTime
            If fdate >= oDate Then
                lastBillGen = fdate
            Else
                lastBillGen = oDate
            End If
            If firstrec = True Then
                Print(fnum, GetStringToPrint(17, "Godown Type", "S") & GetStringToPrint(17, "Godown Number", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(13, "Rent", "N") & GetStringToPrint(13, "Pe Rent", "N") & GetStringToPrint(13, "GST Amt", "N") & GetStringToPrint(13, "Outstanding", "N") & GetStringToPrint(20, "   From Month-Year", "S") & GetStringToPrint(13, " Advance", "S") & vbNewLine)
                Print(fnum, StrDup(165, "-") & vbNewLine)
                firstrec = False
            End If
            chkrs4.Open("SELECT * FROM RENT WHERE [GROUP]='" & grp & "' and GODWN_NO='" & gdn & "' and P_CODE ='" & pcd & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
            Dim ramt As Double = 0
            Dim pramt As Double = 0
            If chkrs4.EOF = False Then
                chkrs4.MoveFirst()
                ramt = chkrs4.Fields(4).Value
                If IsDBNull(chkrs4.Fields(5).Value) Then
                Else
                    pramt = chkrs4.Fields(5).Value
                End If
            End If
            chkrs4.Close()
            Dim gst_amt As Double = 0
            Dim rnd As Integer

            If pgst.Equals("997212") Then
                gst_amt = (ramt + pramt) * 18 / 100
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If
            End If
            monthamt = ramt + pramt + gst_amt
            Dim adv As String = ""
            If chkrs1.EOF = True Then
                Exit Do
            End If
            chkrs3.Open("SELECT [bill].* from [BILL] WHERE [GROUP]='" & grp & "' AND [GODWN_NO]='" & gdn & "' AND BILL_DATE>" & fdate & " order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", xcon)
            While Not chkrs3.EOF And chkrs3.Fields(1).Value = grp And chkrs3.Fields(2).Value = gdn And chkrs3.Fields(3).Value = pcd
                pgross = pgross + chkrs3.Fields(5).Value
                pcgst = pcgst + chkrs3.Fields(7).Value
                psgst = psgst + chkrs3.Fields(9).Value
                pnet = pnet + chkrs3.Fields(10).Value
                lastBillGen = chkrs3.Fields(4).Value
                If chkrs3.EOF = False Then
                    chkrs3.MoveNext()
                End If
                If chkrs3.EOF = True Then
                    Exit While
                End If
            End While
            chkrs3.Close()
            chkrs2.Open("SELECT [receipt].* from [receipt] WHERE [GROUP]='" & grp & "' AND [GODWN_NO]='" & gdn & "' AND [receipt].REC_DATE>" & fdate & " order by [receipt].[GROUP],[receipt].GODWN_NO,[receipt].REC_DATE,[receipt].REC_NO", xcon)
            While Not chkrs2.EOF
                If (chkrs2.Fields(1).Value = grp And chkrs2.Fields(2).Value = gdn) Then
                    rnet = rnet + chkrs2.Fields(5).Value
                End If
                If chkrs2.EOF = False Then
                    chkrs2.MoveNext()
                End If
                If chkrs2.EOF = True Then
                    Exit While
                End If
            End While
            out = pnet - rnet
            If out < 0 Then
                adv = "Advance"
                out = out * -1
                gtotadv = gtotadv + out
            Else
                gtotamt = gtotamt + out
            End If
            chkrs2.Close()
            Dim monthCount As Integer = 0
            If monthamt > 0 Then
                monthCount = out / monthamt
                monthCount = monthCount - 1
            End If
            If adv.Equals("") Then
                outDate = lastBillGen.AddMonths(-1 * monthCount)
            Else
                outDate = lastBillGen.AddMonths(monthCount + 1)
            End If

            Print(fnum, GetStringToPrint(17, grp, "S") & GetStringToPrint(17, gdn, "S") & GetStringToPrint(50, pnm, "S") & GetStringToPrint(13, Format(ramt, "0.00"), "N") & GetStringToPrint(13, Format(pramt, "0.00"), "N") & GetStringToPrint(13, Format(gst_amt, "0.00"), "N") & GetStringToPrint(13, Format(out, "0.00"), "N") & GetStringToPrint(20, "   " & outDate.Month.ToString & "-" & outDate.Year.ToString, "S") & GetStringToPrint(13, " " & adv, "S") & vbNewLine)
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If
        Loop
        Print(fnum, StrDup(165, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(123, "Total Outstanding:", "N") & GetStringToPrint(13, Format(gtotamt, "0.00"), "N") & vbNewLine)
        Print(fnum, GetStringToPrint(123, "Total Advance:", "N") & GetStringToPrint(13, Format(gtotadv, "0.00"), "N") & vbNewLine)
        chkrs1.Close()
        xcon.Close()
        FileClose(fnum)

        Form23.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\outstand.dat", RichTextBoxStreamType.PlainText)

        Form23.Show()
        CreatePDF(Application.StartupPath & "\Reports\outstand.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
    End Sub
    Private Function CreatePDF(strReportFilePath As String, invoice_no As String)
        Try
            Dim line As String
            Dim readFile As System.IO.TextReader = New StreamReader(strReportFilePath)
            Dim yPoint As Integer = 60

            Dim pdf As PdfDocument = New PdfDocument
            pdf.Info.Title = "Text File to PDF"

            Dim pdfPage As PdfPage = pdf.AddPage

            ' pdfPage.Orientation = PdfSharp.PageOrientation.Landscape
            pdfPage.TrimMargins.Left = 15

            pdfPage.Width = 842
            pdfPage.Height = 595

            '  pdf.Pages.RemoveAt(0)
            Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
            Dim font As XFont = New XFont("COURIER NEW", 7, XFontStyle.Regular)

            Dim counter As Integer
            While True
                counter = counter + 1

                line = readFile.ReadLine()

                If line Is Nothing Then
                    Exit While
                Else
                    If counter > 43 Then
                        counter = 1
                        pdfPage = pdf.AddPage()
                        graph = XGraphics.FromPdfPage(pdfPage)
                        font = New XFont("COURIER NEW", 7, XFontStyle.Regular)

                        pdfPage.TrimMargins.Left = 15

                        pdfPage.Width = 842
                        pdfPage.Height = 595
                        yPoint = 60
                    End If
                    'If counter = 1 Or counter = 31 Then
                    '    font = New XFont("COURIER NEW", 14, XFontStyle.Bold)
                    'Else
                    '    font = New XFont("COURIER NEW", 10, XFontStyle.Regular)
                    'End If
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

    Private Sub FrmOutstanding_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.KeyPreview = True
        RadioGodown.Checked = True
        formloaded = True
    End Sub

    Private Sub FrmOutstanding_Move(sender As Object, e As EventArgs) Handles Me.Move
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        fnum = FreeFile() '''''''''Get FreeFile No.'''''''''''
        Dim numRec As Integer = 0
        Dim xline As Integer = 0
        FileOpen(fnum, Application.StartupPath & "\Reports\outstand.dat", OpenMode.Output)
        If xcon.State = ConnectionState.Open Then
        Else
            xcon.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\millrc.accdb;")
        End If
        If RadioGodown.Checked = True Then
            chkrs1.Open("SELECT [godown].*,[PARTY].P_NAME from [godown] INNER JOIN [PARTY] ON [godown].P_CODE=[PARTY].P_CODE WHERE [STATUS]='C' order by [godown].[GROUP],[godown].GODWN_NO,[godown].P_CODE", xcon)
        Else
            chkrs1.Open("SELECT [godown].*,[PARTY].P_NAME from [godown] INNER JOIN [PARTY] ON [godown].P_CODE=[PARTY].P_CODE WHERE [STATUS]='C' order by [PARTY].P_NAME", xcon)
        End If
        Dim gtotamt As Double = 0
        Dim gtotadv As Double = 0
        Dim firstrec As Boolean = True
        Do While chkrs1.EOF = False
            Dim grp As String = chkrs1.Fields(0).Value
            Dim gdn As String = chkrs1.Fields(3).Value
            Dim pcd As String = chkrs1.Fields(1).Value
            Dim pnm As String = chkrs1.Fields(38).Value
            Dim fdate As String = chkrs1.Fields(11).Value
            Dim pgst As String = chkrs1.Fields(37).Value
            Dim pgross As Double = 0
            Dim pcgst As Double = 0
            Dim psgst As Double = 0
            Dim pnet As Double = 0
            Dim rnet As Double = 0
            Dim out As Double = 0
            Dim lastBillGen As Date
            Dim monthamt As Double
            Dim iDate As String = "31/07/2017"
            Dim oDate As DateTime = Convert.ToDateTime(iDate)
            Dim outDate As DateTime
            If fdate >= oDate Then
                lastBillGen = fdate
            Else
                lastBillGen = oDate
            End If
            If firstrec = True Then
                Print(fnum, GetStringToPrint(17, "Godown Type", "S") & GetStringToPrint(17, "Godown Number", "S") & GetStringToPrint(50, "Tenant Name", "S") & GetStringToPrint(13, "Rent", "N") & GetStringToPrint(13, "Pe Rent", "N") & GetStringToPrint(13, "GST Amt", "N") & GetStringToPrint(13, "Outstanding", "N") & GetStringToPrint(20, "   From Month-Year", "S") & GetStringToPrint(13, " Advance", "S") & vbNewLine)
                Print(fnum, StrDup(165, "-") & vbNewLine)
                firstrec = False
            End If
            chkrs4.Open("SELECT * FROM RENT WHERE [GROUP]='" & grp & "' and GODWN_NO='" & gdn & "' and P_CODE ='" & pcd & "' order by  DateValue('01/'+STR(FR_MONTH)+'/'+STR(FR_YEAR)) DESC", xcon)
            Dim ramt As Double = 0
            Dim pramt As Double = 0
            If chkrs4.EOF = False Then
                chkrs4.MoveFirst()
                ramt = chkrs4.Fields(4).Value
                If IsDBNull(chkrs4.Fields(5).Value) Then
                Else
                    pramt = chkrs4.Fields(5).Value
                End If
            End If
            chkrs4.Close()
            Dim gst_amt As Double = 0
            Dim rnd As Integer

            If pgst.Equals("997212") Then
                gst_amt = (ramt + pramt) * 18 / 100
                rnd = gst_amt - Math.Round(gst_amt)
                If rnd >= 50 Then
                    gst_amt = Math.Round(gst_amt) + 1
                Else
                    gst_amt = Math.Round(gst_amt)
                End If
            End If
            monthamt = ramt + pramt + gst_amt
            Dim adv As String = ""
            If chkrs1.EOF = True Then
                Exit Do
            End If
            chkrs3.Open("SELECT [bill].* from [BILL] WHERE [GROUP]='" & grp & "' AND [GODWN_NO]='" & gdn & "' AND BILL_DATE>" & fdate & " order by [BILL].[GROUP],[BILL].GODWN_NO,[BILL].BILL_DATE,[BILL].INVOICE_NO", xcon)
            While Not chkrs3.EOF And chkrs3.Fields(1).Value = grp And chkrs3.Fields(2).Value = gdn And chkrs3.Fields(3).Value = pcd
                pgross = pgross + chkrs3.Fields(5).Value
                pcgst = pcgst + chkrs3.Fields(7).Value
                psgst = psgst + chkrs3.Fields(9).Value
                pnet = pnet + chkrs3.Fields(10).Value
                lastBillGen = chkrs3.Fields(4).Value
                If chkrs3.EOF = False Then
                    chkrs3.MoveNext()
                End If
                If chkrs3.EOF = True Then
                    Exit While
                End If
            End While
            chkrs3.Close()
            chkrs2.Open("SELECT [receipt].* from [receipt] WHERE [GROUP]='" & grp & "' AND [GODWN_NO]='" & gdn & "' AND [receipt].REC_DATE>" & fdate & " order by [receipt].[GROUP],[receipt].GODWN_NO,[receipt].REC_DATE,[receipt].REC_NO", xcon)
            While Not chkrs2.EOF
                If (chkrs2.Fields(1).Value = grp And chkrs2.Fields(2).Value = gdn) Then
                    rnet = rnet + chkrs2.Fields(5).Value
                End If
                If chkrs2.EOF = False Then
                    chkrs2.MoveNext()
                End If
                If chkrs2.EOF = True Then
                    Exit While
                End If
            End While
            out = pnet - rnet
            If out < 0 Then
                adv = "Advance"
                out = out * -1
            End If
            chkrs2.Close()
            Dim monthCount As Integer = 0
            If monthamt > 0 Then
                monthCount = out / monthamt
                monthCount = monthCount - 1
            End If
            If adv.Equals("") Then
                outDate = lastBillGen.AddMonths(-1 * monthCount)
            Else
                outDate = lastBillGen.AddMonths(monthCount + 1)
            End If

            Print(fnum, GetStringToPrint(17, grp, "S") & GetStringToPrint(17, gdn, "S") & GetStringToPrint(50, pnm, "S") & GetStringToPrint(13, Format(ramt, "0.00"), "N") & GetStringToPrint(13, Format(pramt, "0.00"), "N") & GetStringToPrint(13, Format(gst_amt, "0.00"), "N") & GetStringToPrint(13, Format(out, "0.00"), "N") & GetStringToPrint(20, "   " & outDate.Month.ToString & "-" & outDate.Year.ToString, "S") & GetStringToPrint(13, " " & adv, "S") & vbNewLine)
            If chkrs1.EOF = False Then
                chkrs1.MoveNext()
            End If
        Loop
        Print(fnum, StrDup(165, "-") & vbNewLine)
        Print(fnum, GetStringToPrint(123, "Total Outstanding:", "N") & GetStringToPrint(13, Format(gtotamt, "0.00"), "N") & vbNewLine)
        Print(fnum, GetStringToPrint(123, "Total Advance:", "N") & GetStringToPrint(13, Format(gtotadv, "0.00"), "N") & vbNewLine)
        chkrs1.Close()
        xcon.Close()
        FileClose(fnum)

        Form23.RichTextBox1.LoadFile(Application.StartupPath & "\Reports\outstand.dat", RichTextBoxStreamType.PlainText)

        Form23.Show()
        CreatePDF(Application.StartupPath & "\Reports\outstand.dat", Application.StartupPath & "\Reports\" & TextBox5.Text)
        Dim PrintPDFFile As New ProcessStartInfo
        PrintPDFFile.UseShellExecute = True

        PrintPDFFile.Verb = "print"
        PrintPDFFile.WindowStyle = ProcessWindowStyle.Normal
        '.Hidden
        PrintPDFFile.FileName = Application.StartupPath & "\Reports\" & TextBox5.Text & ".pdf"
        Process.Start(PrintPDFFile)
    End Sub
End Class