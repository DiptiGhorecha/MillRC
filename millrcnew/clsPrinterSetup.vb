Public Class clsPrinterSetup
    '***************************************************************
    '******* Purpose :  To Set Paper Size,Lines Per Page and *******
    '*******            Printer Control Language Settings    *******
    '*******            For Printing With Win32 API          *******
    '*******            Table : PrinterSetup in Company.mdb  *******
    '***************************************************************

    'Private Objects for this Class
    Private mPrinterID As Integer
    Private mPrinterType As String
    Private mStartBold As String
    Private mEndBold As String
    Private mPageLength As String
    Private mSingleCompress As String   'For Draft 12cpi
    Private mDblCompress As String      'For Draft 17cpi
    Private mDblToDblCompress As String 'For Draft 20cpi
    Private mCancelCompress As String
    Private mResetPrinter As String
    Private mFormFeed As String
    Private mblnSelected As Boolean
    Private mNornalFont As String   'For Draft 10cpi
    Private mLinesPerPage As Integer   'For Lines per Page
    Private mPageSize As String   'For Page Size - A4 Size and 80 Cols
    Dim vdatai As Integer = 0
    Dim vdata As String = ""

    Public Property PrinterID As Integer
        Set
            mPrinterID = vData
        End Set
        Get
            PrinterID = mPrinterID
        End Get

    End Property

    Public Property StartBold As String
        Set
            mStartBold = vdata
        End Set
        Get
            StartBold = mStartBold
        End Get

    End Property


    Public Property PrinterType As String
        Set
            mPrinterType = vdata
        End Set
        Get
            PrinterType = mPrinterType
        End Get

    End Property
    Public Property EndBold As String
        Set
            mEndBold = vdata
        End Set
        Get
            EndBold = mEndBold
        End Get

    End Property
    Public Property PageLength As String
        Set
            mPageLength = vdata
        End Set
        Get
            PageLength = mPageLength
        End Get

    End Property
    Public Property NormalFont As String
        Set
            mNornalFont = vdata
        End Set
        Get
            NormalFont = mNornalFont
        End Get

    End Property
    Public Property SingleCompress As String
        Set
            mSingleCompress = vdata
        End Set
        Get
            SingleCompress = mSingleCompress
        End Get

    End Property
    Public Property DoubleCompress As String
        Set
            mDblCompress = vdata
        End Set
        Get
            DoubleCompress = mDblCompress
        End Get

    End Property
    Public Property DblToDblCompress As String
        Set
            mDblToDblCompress = vdata
        End Set
        Get
            DblToDblCompress = mDblToDblCompress
        End Get

    End Property
    Public Property CancelCompress As String
        Set
            mCancelCompress = vdata
        End Set
        Get
            CancelCompress = mCancelCompress
        End Get

    End Property
    Public Property ResetPrinter As String
        Set
            mResetPrinter = vdata
        End Set
        Get
            ResetPrinter = mResetPrinter
        End Get

    End Property
    Public Property FormFeed As String
        Set
            mFormFeed = vdata
        End Set
        Get
            FormFeed = mFormFeed
        End Get

    End Property
    Public Property LinesPerPage As String
        Set
            mLinesPerPage = vdata
        End Set
        Get
            LinesPerPage = mLinesPerPage
        End Get

    End Property
    Public Property PageSize As String
        Set
            mPageSize = vdata
        End Set
        Get
            PageSize = mPageSize
        End Get

    End Property
    Public Property IsSelected As String
        Set
            mblnSelected = vdata
        End Set
        Get
            IsSelected = mblnSelected
        End Get

    End Property

    Public Function SavePrinterSettings(ByVal vID As Integer) As Boolean

        On Error GoTo Errhandle
        SavePrinterSettings = False

        Dim mConnection As ADODB.Connection
        Dim mRsPrinter As ADODB.Recordset

        mConnection = New ADODB.Connection
        mRsPrinter = New ADODB.Recordset
        mConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\RAS.mdb;Jet OLEDB:Database Password=ras258;")
        mRsPrinter.Open("Select * from PrinterSetup where PrinterID=" & vID, mConnection)

        mConnection.Execute("Update PrinterSetup set blnSelected=false ")

        If mRsPrinter.EOF Then Exit Function

        With mRsPrinter
            !PrinterID = mPrinterID
            !PrinterType = mPrinterType
            !StartBold = mStartBold
            !EndBold = mEndBold
            !PageLength = mPageLength
            !NormalFont = mNornalFont
            !SingleCompress = mSingleCompress
            !DblCompress = mDblCompress
            !DblToDblCompress = mDblToDblCompress
            !CancelCompress = mCancelCompress
            !ResetPrinter = mResetPrinter
            !FormFeed = mFormFeed
            !blnSelected = mblnSelected
            !PageSize = mPageSize
            !LinesPerPage = mLinesPerPage
            .Update()
        End With
        SavePrinterSettings = True
        mRsPrinter = Nothing
        mConnection = Nothing

        Exit Function
Errhandle:
        MsgBox("Printer Setup Class : " & Err.Description)

    End Function

    Public Function GetPrinterSettings(ByVal vID As Integer, vDefault As Boolean) As Boolean

        Dim mConnection As ADODB.Connection
        Dim mRsPrinter As ADODB.Recordset

        mConnection = New ADODB.Connection
        mRsPrinter = New ADODB.Recordset
        'xcon.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\MH_Share.mdb"
        mConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\RAS.mdb;Jet OLEDB:Database Password=ras258;")
        If vDefault = False Then
            mRsPrinter.Open("Select * from PrinterSetup where PrinterID=" & vID, mConnection)
        Else
            mRsPrinter.Open("Select * from PrinterSetup where blnSelected=True", mConnection)
        End If

        With mRsPrinter

            If .RecordCount > 0 Then
                mPrinterID = !PrinterID
                mPrinterType = !PrinterType & ""
                mStartBold = !StartBold & ""
                mEndBold = !EndBold & ""
                mPageLength = !PageLength & ""
                mNornalFont = !NormalFont & ""
                mSingleCompress = !SingleCompress & ""
                mDblCompress = !DblCompress & ""
                mDblToDblCompress = !DblToDblCompress & ""
                mCancelCompress = !CancelCompress & ""
                mResetPrinter = !ResetPrinter & ""
                mblnSelected = !blnSelected & ""
                mFormFeed = !FormFeed & ""
                mPageSize = !PageSize & ""
                mLinesPerPage = IIf(IsDBNull(!LinesPerPage), 0, !LinesPerPage)

            End If
        End With

        mRsPrinter = Nothing
        mConnection = Nothing
    End Function
    Public Sub FillPrinterType(ByVal vObj As ComboBox)

        Dim mConnection As ADODB.Connection
        Dim mRsPrinter As ADODB.Recordset

        mConnection = New ADODB.Connection
        mRsPrinter = New ADODB.Recordset

        mConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\RAS.mdb;Jet OLEDB:Database Password=ras258;")
        mRsPrinter.Open("Select * from PrinterSetup ", mConnection)
        Dim lItem As clsItem
        Dim x As Integer
        With mRsPrinter
            .MoveFirst()

            Do Until .EOF
                lItem = New clsItem()
                lItem.DisplayText = !PrinterType
                lItem.ItemData = !PrinterID

                x = vObj.Items.Add(lItem)

                If !blnSelected = True Then vObj.SelectedIndex = x
                .MoveNext()
            Loop
        End With

        mRsPrinter = Nothing
        mConnection = Nothing
    End Sub


    Private Sub Class_Initialize()
        GetPrinterSettings(0, True)
    End Sub

End Class
Public Class clsItem
    Dim ItemText As String
    Dim ItemID As String

    Public Property DisplayText() As String
        Get
            DisplayText = ItemText
        End Get
        Set(ByVal Value As String)
            ItemText = Value
        End Set
    End Property

    Public Property ItemData() As String
        Get
            ItemData = ItemID
        End Get
        Set(ByVal Value As String)
            ItemID = Value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return ItemText
    End Function

End Class
