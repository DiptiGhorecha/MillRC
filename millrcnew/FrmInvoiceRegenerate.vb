Option Explicit On
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
''' <summary>
''' tables used - bill
''' this module was created to place logo on all invoice pdf files. Now logo checkbox is place on invoice generate module, so we can hide/remove this module.
''' </summary>
Public Class FrmInvoiceRegenerate
    Dim chkrs1 As New ADODB.Recordset
    Dim chkrs2 As New ADODB.Recordset
    Dim xcon As New ADODB.Connection    ''''''''variable is used to open a connection
    Dim pdfpath As String
    Dim strReportFilePath As String
    Dim formloaded As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
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

            chkrs1.Open("SELECT * FROM BILL ORDER BY BILL_DATE,INVOICE_NO", xcon)
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = chkrs2.Fields(0).Value
            chkrs2.Close()
            ProgressBar1.Value = 0
            Dim counter As Integer = 0
            Do While chkrs1.EOF = False
                counter = counter + 1
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

                pdfpath = Application.StartupPath & "\Invoices\pdf\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month)
                strReportFilePath = Application.StartupPath & "\Invoices\dat\" & BLDT.Year.ToString & "\" & MonthName(BLDT.Month) & "\" & INVNO & ".dat"
                Console.WriteLine(INVNO)
                CreatePDF(strReportFilePath, INVNO)
                ProgressBar1.Value += 1

                If chkrs1.EOF = False Then
                    chkrs1.MoveNext()
                End If
            Loop
            chkrs1.Close()

            xcon.Close()

            MessageBox.Show("Process completed......")
        Catch ex As Exception
            MsgBox("exception1 --> " & ex.ToString)
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
                    graph.DrawString(line, font, XBrushes.Black,
                     New XRect(50, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                    yPoint = yPoint + 12
                End If
            End While
            Dim pdfFilename As String = pdfpath & "\" & invoice_no & ".pdf"
            pdf.Save(pdfFilename)
            readFile.Close()
            readFile.Dispose()
            readFile = Nothing
            pdf.Close()
            pdf.Dispose()
        Catch ex As Exception
            MsgBox("Exception2 --> " & ex.ToString)
        End Try
    End Function

    Private Sub FrmInvoiceRegenerate_Load(sender As Object, e As EventArgs) Handles Me.Load
        '''''''set position of the form
        Try
            Me.MdiParent = MainMDIForm
            Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
            Me.Left = 0
            Dim MyDate As Date = Now
            Dim DaysInMonth As Integer
            formloaded = True
        Catch ex As Exception
            MessageBox.Show("Error opening bulk invoice regenerate: " & ex.Message)
        End Try
    End Sub

    Private Sub FrmInvoiceRegenerate_Move(sender As Object, e As EventArgs) Handles Me.Move
        ''''don't allow user to move the form
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Me.Close()    ''''close form
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Error Cancel Module: " & ex.Message)
        End Try
    End Sub


End Class