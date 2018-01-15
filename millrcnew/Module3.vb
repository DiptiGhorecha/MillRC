Module Module3
    Public datadrive As String
    Public xdatabase As String
    Public newname As String
    Public bankrptt As Integer
    Public prnmaxpagelines As Integer
    Public QFMT, VFMT, VFMT2, DFMT, VFMTDb As String
    Public IMAGEFLNAME As String
    Public XCOMPANYNAME As String
    Public XAddress As String
    Public XRptTitle As String
    Public prnleftmargin As Integer
    Public one(19)
    Public ty(9)
    Public MODE As String
    Public hun(5)
    Public figwordnum As String
    Public figwordword As String
    Public Const PRNSeparator = " "
    Public Const PRNA4Paper = "A4"
    Public Const PRN80ColPaper = "80"
    Public Const prn132colpaper = "132"
    Public DBName As DAO.Database
    Public Declare Function GetDesktopWindow Lib "USER32" () As Long
    Public strBlankglb As Integer
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long,
ByVal lpOperation As String,
ByVal lpFile As String,
ByVal lpParameters As String,
ByVal lpDirectory As String,
ByVal nShowCmd As Long) As Long

    'for File Copy Animation and Copy to Recycle bin

    Public Const FO_MOVE As Long = &H1
    Public Const FO_COPY As Long = &H2
    Public Const FO_DELETE As Long = &H3
    Public Const FO_RENAME As Long = &H4

    Public Const FOF_MULTIDESTFILES As Long = &H1
    Public Const FOF_CONFIRMMOUSE As Long = &H2
    Public Const FOF_SILENT As Long = &H4
    Public Const FOF_RENAMEONCOLLISION As Long = &H8
    Public Const FOF_NOCONFIRMATION As Long = &H10
    Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
    Public Const FOF_CREATEPROGRESSDLG As Long = &H0
    Public Const FOF_ALLOWUNDO As Long = &H40
    Public Const FOF_FILESONLY As Long = &H80
    Public Const FOF_SIMPLEPROGRESS As Long = &H100
    Public Const FOF_NOCONFIRMMKDIR As Long = &H200
    Public Structure SHFILEOPSTRUCT

        Dim hWnd As Long
        Dim wFunc As Long
        Dim pFrom As String
        Dim pTo As String
        Dim fFlags As Integer
        Dim fAnyOperationsAborted As Boolean
        Dim hNameMappings As Long
        Dim lpszProgressTitle As String
    End Structure
    Public Structure BrowseInfo
        Dim hWndOwner As Long
        Dim pidlRoot As Long
        Dim sDisplayName As String
        Dim sTitle As String
        Dim ulFlags As Long
        Dim lpfn As Long
        Dim lParam As Long
        Dim iImage As Long
    End Structure
    Public Declare Function SHFileOperation Lib "shell32.dll" Alias _
  "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
  (bBrowse As BrowseInfo) As Long
    Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  (ByVal lItem As Long, ByVal sDir As String) As Long
    Public Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Public screenWidth As Long, screenHeight As Long


    '------------------------------------------------------------------
    'Class for Printer Settings
    Public objPRNSetup As clsPrinterSetup


    Public Sub Main()
        DFMT = "dd-MM-yyyy"
        datadrive = CurDir() 'later on this will be it
        If Len(Application.StartupPath) > 3 Then
            datadrive = Application.StartupPath & "\"
        Else
            datadrive = Application.StartupPath
        End If
        xdatabase = "MH_Share.MDB"
    End Sub

    Public Function FileExists(ByVal fileName$) As Boolean
        Dim X As Long
        On Error GoTo FileError
        X = FileLen(fileName$)
        FileExists = True
        Exit Function
FileError:
        FileExists = False
        Exit Function
    End Function


    '-------------------------------------------------------------
    'Purpose :- To Get String for Printing
    '-------------------------------------------------------------
    Public Function GetStringToPrint(ByVal iWidth As Integer, ByVal strData As String, ByVal vtype As String) As String
        Dim strBlank As String
        strData = Left(strData, iWidth)
        strBlank = Space(iWidth - Len(strData))
        strBlankglb = iWidth - Len(strData)
        Select Case vtype
            Case "S"    'for String
                GetStringToPrint = strData & strBlank
            Case "N"    'for Numeric
                GetStringToPrint = strBlank & strData
            Case "C"    'For Centre
                strBlank = Space(Math.Round(Len(strBlank) / 2, 0))
                GetStringToPrint = strBlank & strData
                GetStringToPrint = GetStringToPrint & Space(iWidth - Len(GetStringToPrint))
        End Select
    End Function
    Public Sub Cent(vLen As Integer, vReportTitle As String, vFreeFile As Integer, vMode As String)

        On Error GoTo errorhandlor  '   Error handlor
        'printer.FontName = "Draft 6Cpi"
        'printer.CurrentX = (printer.Width - printer.TextWidth(UCase(xcompanyname))) / 2
        'printer.Print UCase(xcompanyname)
        Dim strTitle As String
        XCOMPANYNAME = "The Motilal Hirabhai Estate and Warehouse Ltd"

        If vMode <> "S" Then
            strTitle = Space((vLen - Len(XCOMPANYNAME)) / 2) & UCase(XCOMPANYNAME)
            Print(vFreeFile, Chr(27) & "G" & Space(prnleftmargin) & strTitle & Chr(27) & "H")
        Else
            strTitle = Space((vLen - Len(XCOMPANYNAME)) / 2) & UCase(XCOMPANYNAME)
            Print(vFreeFile, Space(prnleftmargin) & strTitle)
        End If


        'printer.Font = "Draft 10Cpi"
        'printer.CurrentX = (printer.Width - printer.TextWidth(UCase(frmname))) / 2
        'printer.Print UCase(frmname)
        'strTitle = Space((vLen - Len(vReportTitle)) / 2) & UCase(vReportTitle)
        strTitle = GetStringToPrint(vLen, UCase(vReportTitle), "C")
        Print(vFreeFile, Space(prnleftmargin) & strTitle)
        Exit Sub
errorhandlor:
        '   If error then catch it and display appropriate message
        If Err.Number = "482" Then
            MsgBox("Printer Not Ready... Please try again")
        ElseIf Err.Number = 91 Then

        End If

    End Sub
    Public Function DoEval(ByVal vString As String) As String
        Dim mString() As String
        'Dim sc1 As ScriptControl
        Dim i As Integer
        Dim strPRNCode As String

        mString = Split(vString, "+")
        'Split Printer Control code by "+"
        For i = 0 To UBound(mString)
            strPRNCode = Replace(UCase(mString(i)), "CHR", "")
            strPRNCode = Replace(UCase(strPRNCode), "$", "")
            strPRNCode = Replace(UCase(strPRNCode), "(", "")
            strPRNCode = Replace(UCase(strPRNCode), ")", "")
            DoEval = DoEval + Strings.Chr(Trim(strPRNCode))
        Next

    End Function
    '---------------------------------
    'Purpose : To Get Amount in word
    '---------------------------------

    Public Function FigInWord(ByVal strAmount As String) As String 'Command1_Click()
        Dim nosplit(5)
        Dim resplit(5)
        Dim p As Integer
        Dim nos, str1, pais As String
        Dim i As Integer
        Dim result As String
        one(1) = "One "
        one(2) = " Two "
        one(3) = " Three "
        one(4) = " Four "
        one(5) = " Five "
        one(6) = " Six "
        one(7) = " Seven "
        one(8) = " Eight "
        one(9) = " Nine "
        one(10) = " Ten "
        one(11) = "Eleven "
        one(12) = "Twelve "
        one(13) = "Thirteen "
        one(14) = "Fourteen "
        one(15) = "Fifteen "
        one(16) = "Sixteen "
        one(17) = "Seventeen "
        one(18) = "Eighteen "
        one(19) = "Ninteen "

        ty(1) = ""
        ty(2) = "Twenty "
        ty(3) = "Thirty "
        ty(4) = "Fourty "
        ty(5) = "Fifty "
        ty(6) = "Sixty "
        ty(7) = "Seventy "
        ty(8) = "Eighty "
        ty(9) = "Ninety "


        hun(1) = "Crore(s) "
        hun(2) = "Lakh(s) "
        hun(3) = "Thousand "
        hun(4) = "Hundred "

        'This Following statement is used for specify the input length 9
        nos = Trim(strAmount) 'figwordnum  'Trim(inno.Text)
        p = InStr(1, nos, ".", 1)
        pais = Mid(nos, p + 1, 2)

        If Len(pais) = 1 Then
            pais = pais + "0"
        End If

        If p > 0 Then
            nos = Mid(nos, 1, p - 1)
        End If

        'nos = String((9 - Len(nos)), "0") + nos
        nos = New String("0"c, (9 - Len(nos))) + nos
        'split statement
        nosplit(1) = Val(Mid(nos, 1, 2))
        nosplit(2) = Val(Mid(nos, 3, 2))
        nosplit(3) = Val(Mid(nos, 5, 2))
        nosplit(4) = Val(Mid(nos, 7, 1))
        nosplit(5) = Val(Mid(nos, 8, 2))
        Dim spli, spli1, spli2 As Integer

        For i = 1 To 5
            spli = nosplit(i)
            If spli > 0 And spli < 20 Then
                resplit(i) = Trim(one(spli)) + " "
            End If

            If spli > 19 Then
                spli1 = Val(Mid(Trim(spli), 1, 1))
                spli2 = Val(Mid(Trim(spli), 2, 1))
                resplit(i) = Trim(ty(spli1)) + " "
                If spli2 > 0 Then
                    resplit(i) = Trim(ty(spli1)) + " " + Trim(one(spli2)) + " "
                End If
            End If

            If Not resplit(i) = "" Then
                result = result & resplit(i) & hun(i)
            End If
        Next i
        Dim paise As String
        Dim pais1, pais2 As Integer
        'paise calculations
        If p > 0 Then
            If Val(pais) > 0 And Val(pais) < 20 Then
                paise = Trim(one(pais)) + " "
            End If

            If pais > 19 Then
                pais1 = Val(Mid(Trim(pais), 1, 1))
                pais2 = Val(Mid(Trim(pais), 2, 1))
                paise = Trim(ty(pais1)) + " "
                If pais2 > 0 Then
                    paise = Trim(ty(pais1)) + " " + Trim(one(pais2))
                End If
            End If
        Else
            paise = ""
        End If

        If p > 0 Then
            FigInWord = result + " And " + paise + " Paise Only"
        Else
            FigInWord = result + " Only"
        End If

    End Function

    Public Sub WriteFile(ByVal xTxtPth As String, ByVal xDir As String)
        Dim intCtr As Integer   ' Loop Counter
        Dim intFNum As Integer  ' File Number
        Dim intMsg As Integer   ' For Msgbox ()

        intFNum = FreeFile()
        FileOpen(intFNum, xTxtPth & "\CircularView.Dat", OpenMode.Output)
        PrintLine(intFNum, xDir)
        FileClose(intFNum)
    End Sub
    Public Sub WriteFileProxy(ByVal xTxtPth As String, ByVal xDir As String)
        Dim intCtr As Integer   ' Loop Counter
        Dim intFNum As Integer  ' File Number
        Dim intMsg As Integer   ' For Msgbox ()
        intFNum = FreeFile()
        FileOpen(intFNum, xTxtPth & "\ProxyForm.Dat", OpenMode.Output)
        PrintLine(intFNum, xDir)
        FileClose(intFNum)
    End Sub
    Public Sub Centre(vLen As Integer, vReportTitle As String, vFreeFile As Integer, vMode As String)

        On Error GoTo errorhandlor  '   Error handlor
        'printer.FontName = "Draft 6Cpi"
        'printer.CurrentX = (printer.Width - printer.TextWidth(UCase(xcompanyname))) / 2
        'printer.Print UCase(xcompanyname)
        Dim strTitle As String
        If vMode <> "V" Then
            If Trim(objPRNSetup.StartBold) <> "" Or Trim(objPRNSetup.EndBold) <> "" Then
                strTitle = Space((vLen - Len(XCOMPANYNAME)) / 2) & DoEval(objPRNSetup.StartBold) & XCOMPANYNAME & DoEval(objPRNSetup.EndBold)
            Else
                strTitle = Space((vLen - Len(XCOMPANYNAME)) / 2) & XCOMPANYNAME
            End If
            Print(vFreeFile, strTitle)
        Else
            strTitle = Space((vLen - Len(XCOMPANYNAME)) / 2) & XCOMPANYNAME
            Print(vFreeFile, strTitle)
        End If
        strTitle = Space((vLen - Len(XAddress)) / 2) & XAddress
        Print(vFreeFile, strTitle)
        'printer.Font = "Draft 10Cpi"
        'printer.CurrentX = (printer.Width - printer.TextWidth(UCase(frmname))) / 2
        'printer.Print UCase(frmname)
        'strTitle = Space((vLen - Len(vReportTitle)) / 2) & UCase(vReportTitle)
        strTitle = Space((vLen - Len(vReportTitle)) / 2) & vReportTitle
        'strTitle = GetStringToPrint(vLen, UCase(vReportTitle), "C")
        Print(vFreeFile, strTitle)
        Exit Sub
errorhandlor:
        '   If error then catch it and display appropriate message
        If Err.Number = "482" Then
            MsgBox("Printer Not Ready... Please try again")
        ElseIf Err.Number = 91 Then

        End If

    End Sub
    '------------------------------------------------------------------------------------------------
    'Purpose : procedure for printing the footer part of the report which takes page size as argument
    '------------------------------------------------------------------------------------------------
    Public Sub Set_Footerp(Size As Integer, vFreeFile As Integer)
        On Error GoTo errorhandlor  '   For handling the error
        Dim strLine As String = DrawLine(Size)
        Print(vFreeFile, strLine)     '   print the line
        Print(vFreeFile, Space((Size - 23)) & "Share Accounting System")
errorhandlor:
        '   If error then catch it and displays appropriate message
        If Err.Number = "482" Then
            MsgBox("Printer Not Ready... Please try again")
        End If
    End Sub

    'This Function returns String to Draw Line on Printer
    Public Function DrawLine(ByVal PrinterCol As Integer) As String
        DrawLine = New String("-"c, PrinterCol)
    End Function
    Public Function GetStringXPrint(ByVal iWidth As Integer, ByVal strData As String, ByVal vtype As String) As String
        Dim strBlank As String
        strData = Left(strData, iWidth)
        strBlank = New String("*"c, (iWidth - Len(strData)))
        strBlankglb = iWidth - Len(strData)
        Select Case vtype
            Case "S"    'for String
                GetStringXPrint = strData & strBlank
            Case "N"    'for Numeric
                GetStringXPrint = strBlank & strData
            Case "C"    'For Centre
                strBlank = Space(Math.Round(Len(strBlank) / 2, 0))
                GetStringXPrint = strBlank & strData
                GetStringXPrint = GetStringXPrint & Space(iWidth - Len(GetStringXPrint))
        End Select
    End Function
    Public Sub Set_FooterpTest(Size As Integer, vFreeFile As Integer)
        On Error GoTo errorhandlor  '   For handling the error
        Print(vFreeFile, Space((Size - 23)) & "Share Accounting System")
errorhandlor:
        '   If error then catch it and displays appropriate message
        If Err.Number = "482" Then
            MsgBox("Printer Not Ready... Please try again")
        End If
    End Sub


End Module
