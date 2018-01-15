Public Class clsPrintFile
    '**************************************************************
    '******* Purpose :  Printing of Reportfile              *******
    '*******            Through WIN32 API Method            *******
    '**************************************************************
    'Declaration of API
    Private Structure DOCINFO
        Dim pDocName As String
        Dim pOutputFile As String
        Dim pDatatype As String
    End Structure

    ' Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
    ' hPrinter As Long) As Long
    ' Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
    'hPrinter As Long) As Long
    ' Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
    'hPrinter As Long) As Long
    ' Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
    '"OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long,
    ' ByVal pDefault As Long) As Long
    ' Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
    '"StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long,
    'pDocInfo As DOCINFO) As Long
    ' Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
    'hPrinter As Long) As Long
    ' Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
    'hPrinter As Long, pBuf As Any, ByVal cdBuf As Long,
    'pcWritten As Long) As Long
    Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As IntPtr) As Boolean
    Private Declare Function EndDocPrinter Lib "winspool.drv" (ByRef hPrinter As IntPtr) As Boolean
    Private Declare Function EndPagePrinter Lib "winspool.drv" (ByRef hPrinter As IntPtr) As Boolean
    Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal src As String, ByRef hPrinter As IntPtr, ByVal pd As IntPtr) As Boolean
    Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByRef hPrinter As IntPtr, ByRef Level As Integer, ByRef pDocInfo As DOCINFO) As Integer
    Private Declare Function StartPagePrinter Lib "winspool.drv" (ByRef hPrinter As IntPtr) As Boolean
    Private Declare Function WritePrinter Lib "winspool.drv" (ByRef hPrinter As IntPtr, pBuf As String, ByRef cdBuf As Integer, ByRef pcWritten As Integer) As Boolean
    Dim NumLoops As Long     ' number of 8k loops
    Dim LeftOver As Integer  ' amount of file left
    Dim i As Integer         ' counter for loops
    Const MaxSize = 8192     ' maximum buffer size
    '------------------------------------------------------------------

    '--------------------------------------------------------------
    'Purpose : To Print Report File
    '--------------------------------------------------------------
    Public Sub PrintReportFile(ByVal strFile As String)

        Dim fnum As Integer
        Dim TextLine
        Dim lhPrinter As Long
        Dim lReturn As Long
        Dim lpcWritten As Long
        Dim lDoc As Long
        Dim sWrittenData As String
        Dim MyDocInfo As DOCINFO
        Dim IsFileOpen As Boolean

        On Error GoTo Errhandle
        '---------------------------------------------------------------------
        'Method - 1     'Print With WIN32 API
        '---------------------------------------------------------------------
        Dim oPS As New System.Drawing.Printing.PrinterSettings
        Dim DefaultPrinterName As String = oPS.PrinterName
        lReturn = OpenPrinter(DefaultPrinterName, lhPrinter, 0)
        If lReturn = 0 Then
            MsgBox("The Printer Name you selected wasn't recognized.")
            Exit Sub
        End If
        MyDocInfo.pDocName = GetFileName(strFile)
        MyDocInfo.pOutputFile = vbNullString
        MyDocInfo.pDatatype = vbNullString
        lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
        Call StartPagePrinter(lhPrinter)
        'Open file for Reading
        fnum = FreeFile()
        'Open strFile For Input As #FNum   ' Open file.
        'Open strFile For Binary Access Read Lock Read As #fNum
        IsFileOpen = True
        '------------------------------------------------
        'Open strFile For Binary As #fnum
        FileOpen(fnum, strFile, OpenMode.Binary)
        'Calculate size of file and amount left over.
        NumLoops = LOF(fnum) \ MaxSize
        LeftOver = LOF(fnum) Mod MaxSize

        'For i = 1 To NumLoops
        sWrittenData = LineInput(fnum)
            lReturn = WritePrinter(lhPrinter, sWrittenData, Len(sWrittenData), lpcWritten)
        ' Next

        ' Grab what is left over.
        'If LeftOver <> 0 Then
        '    sWrittenData = LineInput(fnum)
        '    lReturn = WritePrinter(lhPrinter, sWrittenData,
        '              Len(sWrittenData), lpcWritten)
        'End If

        FileClose(fnum)
        lReturn = EndPagePrinter(lhPrinter)
        lReturn = EndDocPrinter(lhPrinter)
        lReturn = ClosePrinter(lhPrinter)
        '------------------------------------------------
        'Method 2 - Line By Line Printing
        '------------------------------------------------
        'Dim ofs As FileSystemObject
        'Dim oPrinter As TextStream
        ''Set ofs = CreateObject("Scripting.FileSystemObject")
        'Set ofs = New FileSystemObject
        'Set oPrinter = ofs.OpenTextFile(strReportFilePath, ForReading, False)
        'Do While oPrinter.AtEndOfStream <> True
        '
        '    TextLine = oPrinter.ReadLine
        '    If UCase(Left(TextLine, 4)) = "CHR$" Then
        '        sWrittenData = Eval(TextLine)
        '    Else
        '        sWrittenData = TextLine
        '    End If
        '    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        '        Len(sWrittenData), lpcWritten)
        '
        'Loop
        ''Do While Not EOF(fNum)   ' Loop until end of file.
        ''    Line Input #1, TextLine   ' Read line into variable.
        ''   'Debug.Print TextLine   ' Print to the Immediate window.
        ''    'printer.Print TextLine
        ''    sWrittenData = TextLine
        ''    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        ''        Len(sWrittenData), lpcWritten)
        ''Loop
        ''Close #fNum   ' Close file.
        '
        'IsFileOpen = False
        ''Close Printer
        'lReturn = EndPagePrinter(lhPrinter)
        'lReturn = EndDocPrinter(lhPrinter)
        'lReturn = ClosePrinter(lhPrinter)

        '---------------------------------------------------------------------
        'Method -3      'with FSO Object
        '---------------------------------------------------------------------
        'Dim ofs
        'Dim oPrinter
        '  Set ofs = CreateObject("Scripting.FileSystemObject")
        ''  Set oPrinter = ofs.OpenTextFile(strReportFilePath, 2, True, 0)
        ''  oPrinter.WriteLine "This is a test file"
        ''  oPrinter.WriteLine "This is a second line"
        ''  oPrinter.Write Chr(12) ' Formfeed on most printers to eject page
        ''  oPrinter.Close
        '  ofs.CopyFile strReportFilePath, printer.Port, True

        Exit Sub
Errhandle:
        If IsFileOpen = True Then FileClose(fnum)
        MsgBox("PrintFile Class : " & Err.Description)
    End Sub

    '---------------------------------------------------------
    'Purpose: To get ReportFile Name to Show in Printer Spool
    '---------------------------------------------------------
    Private Function GetFileName(ByVal vFullFileName) As String
        Dim strName As String, i As Integer
        For i = Len(vFullFileName) To 1 Step -1
            If Mid(vFullFileName, i, 1) = "\" Then Exit For
            strName = Mid(vFullFileName, i, 1) & strName
        Next
        GetFileName = strName
    End Function


End Class
