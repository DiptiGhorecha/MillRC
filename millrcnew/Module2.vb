Module Module2
    Public SelStart As Long
    Public TextLen As Long
    Public Text As String


    'In a Module
    Public trapUndo As Boolean
    Public UndoStack As New Collection
    Public RedoStack As New Collection

    ' Filter
    Private Structure SHFILEOPTSTRUCT
        Public hWnd As Long
        Public wFunc As Long
        Public pFrom As String
        Public pTo As String
        Public fFlags As Integer
        Public fAnyOperationsAborted As Long
        Public hNameMappings As Long
        Public lpszProgressTitle As Long
    End Structure

    Private Declare Function SHFileOperation Lib "Shell32.dll" _
  Alias "SHFileOperationA" (lpFileOp As SHFILEOPTSTRUCT) As Long

    Private Const FO_DELETE = &H3
    Private Const FOF_ALLOWUNDO = &H40
    Private Declare Function GetVolumeInformation _
        Lib "kernel32" Alias "GetVolumeInformationA" _
        (ByVal lpRootPathName As String,
        ByVal pVolumeNameBuffer As String,
        ByVal nVolumeNameSize As Long,
        lpVolumeSerialNumber As Long,
        lpMaximumComponentLength As Long,
        lpFileSystemFlags As Long,
        ByVal lpFileSystemNameBuffer As String,
        ByVal nFileSystemNameSize As Long) As Long

    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Long) As Long
    Public Const WM_CUT = &H300
    Public Const WM_COPY = &H301
    Public Const WM_PASTE = &H302
    Public Const WM_CLEAR = &H303
    Public Const WM_USER = &H400
    Public Const EM_CANUNDO = &HC6
    Public Const EM_UNDO = &HC7

    Public Const EM_LINEINDEX = &HBB
    Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
    Public Const EM_GETLINECOUNT = &HBA
    Public Const EM_LINEFROMCHAR = &HC9
    ' Win32 Declarations for FolderView
    Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

    ' Win 32 Declarations for View Mode
    Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
    Public Enum ERECViewModes
        ercDefault = 0
        ercWordWrap = 1
        ercWYSIWYG = 2
    End Enum
    Public Sub DeleteFileToRecycleBin(Filename As String)

        Dim fop As SHFILEOPTSTRUCT

        With fop
            .wFunc = FO_DELETE
            .pFrom = Filename
            .fFlags = FOF_ALLOWUNDO

        End With

        SHFileOperation(fop)

    End Sub
    Public Function convinRS(Number As Double) As String
        Dim tempstr As String
        Dim tempnum As Double
        Dim convstr As String
        Dim partlen As Byte
        Dim ptat As Byte
        Dim digs As Integer


        convstr = ""
        ptat = InStr(1, CStr(Number), ".")
        If ptat = 0 Then
            tempnum = CStr(Number)
        Else
            tempnum = Mid(CStr(Number), 1, ptat)
        End If
        tempstr = Trim(CStr(tempnum))

        While Not tempnum = 0

            partlen = Len(tempstr)
            Select Case partlen
                Case Is >= 8
                    digs = CInt(Mid(tempstr, 1, partlen - 7))
                    convstr = diginwrd(digs)
                    If digs > 1 Then
                        convstr = convstr & "Crores "
                    Else
                        convstr = convstr & "Crore "
                    End If
                    tempstr = Right(tempstr, 7)
                Case Is >= 6 And partlen < 8
                    digs = CInt(Mid(tempstr, 1, partlen - 5))
                    convstr = convstr & diginwrd(digs)
                    If digs > 1 Then
                        convstr = convstr & "Lakhs "
                    Else
                        convstr = convstr & "Lakh "
                    End If
                    tempstr = Right(tempstr, 5)

                Case Is < 6 And partlen >= 4
                    digs = CInt(Mid(tempstr, 1, partlen - 3))
                    convstr = convstr + diginwrd(digs)
                    convstr = convstr & "Thousand "
                    tempstr = Right(tempstr, 3)
                Case Is <= 3
                    digs = CInt(tempstr)
                    convstr = convstr + diginwrd(digs)
                    tempstr = "0"
            End Select
            tempnum = CLng(tempstr)

            'MsgBox tempnum & convstr
        End While
        If ptat = 0 Then
            tempnum = 0
        Else
            tempnum = Val(Right(CStr(Number), Len(CStr(Number)) - ptat + 1)) * 100
        End If
        '    MsgBox convstr & tempnum
        'convstr = convstr
        If Not tempnum = 0 Then
            convinRS = convstr & "And " & diginwrd(tempnum) & "Paise Only"
        Else
            convinRS = convstr & "Only"
        End If
    End Function


    Public Function diginwrd(ByVal digsnum As Integer) As String
        Select Case digsnum
            Case 1
                diginwrd = "One "
            Case 2
                diginwrd = "Two "
            Case 3
                diginwrd = "Three "
            Case 4
                diginwrd = "Four "
            Case 5
                diginwrd = "Five "
            Case 6
                diginwrd = "Six "
            Case 7
                diginwrd = "Seven "
            Case 8
                diginwrd = "Eight "
            Case 9
                diginwrd = "Nine "
            Case 10
                diginwrd = "Ten "
            Case 11
                diginwrd = "Eleven "
            Case 12
                diginwrd = "Twelve "
            Case 13
                diginwrd = "Thirteen "
            Case 14
                diginwrd = "Fourteen "
            Case 15
                diginwrd = "Fifteen "
            Case 16
                diginwrd = "Sixteen "
            Case 17
                diginwrd = "Seventeen "
            Case 18
                diginwrd = "Eighteen "
            Case 19
                diginwrd = "Nineteen "
            Case Is > 19
                Dim dig As Integer
                Dim tdigword As String
                Dim thdig As String
                dig = CInt(Right(CStr(digsnum), 1))

                If digsnum >= 100 Then
                    thdig = Left(CStr(digsnum), 1)
                    dig = CInt(Right(CStr(digsnum), 2))
                Else
                    dig = CInt(Right(CStr(digsnum), 1))
                End If
                Select Case digsnum
                    Case Is >= 100
                        tdigword = diginwrd(CInt(thdig)) & "Hundred "
                    Case Is >= 90 And digsnum < 100
                        tdigword = "Ninety "
                    Case Is >= 80 And digsnum < 90
                        tdigword = "Eighty "
                    Case Is >= 70 And digsnum < 80
                        tdigword = "Seventy "
                    Case Is >= 60 And digsnum < 70
                        tdigword = "Sixty "
                    Case Is >= 50 And digsnum < 60
                        tdigword = "Fifty "
                    Case Is >= 40 And digsnum < 50
                        tdigword = "Forty "
                    Case Is >= 30 And digsnum < 40
                        tdigword = "Thirty "
                    Case Is >= 20 And digsnum < 30
                        tdigword = "Twenty "
                End Select
                diginwrd = tdigword & diginwrd(dig)
        End Select

    End Function

End Module
