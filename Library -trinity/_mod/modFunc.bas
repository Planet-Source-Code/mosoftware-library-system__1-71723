Attribute VB_Name = "modFunc"
Option Explicit
Public adoCon As ADODB.Connection
Public adoRes As ADODB.Recordset
'SYSTEM USER DIMENSION
Public iUSER As Integer
Public INT_SIZE As Integer
Public sCon As String
Public sUserId As String

Sub Main()
    sCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\_database\dbLibrary.mdb;Persist Security Info=False"
    frmMain.Show
    frmStartup.Show 1
End Sub

Public Function ENCRYPT(str_encrypt As String) As String
    Dim a As Integer
    Dim b As String
    b = ""
    For a = 1 To Len(str_encrypt)
        '& " "
        If Len(b) > 0 Then b = b & ASCII_TO_BIN(Asc(Mid(str_encrypt, a, 1))) Else b = ASCII_TO_BIN(Asc(Mid(str_encrypt, a, 1)))
    Next
    ENCRYPT = b
End Function

Public Function ASCII_TO_BIN(STR_VAL As Integer) As String
    Dim a(8) As Integer
    Dim b As String
    Dim c As Integer
    a(1) = 1
    a(2) = 2
    a(3) = 4
    a(4) = 8
    a(5) = 16
    a(6) = 32
    a(7) = 64
    a(8) = 128
    c = 8
    Do While c >= 1 And c <= 8
        If a(c) <= STR_VAL Then
            b = b & "1"
            STR_VAL = STR_VAL - a(c)
        Else
            b = b & "0"
        End If
        c = c - 1
    Loop
    ASCII_TO_BIN = b
End Function

Public Sub Set_Icon_btn(cForm As Form, iLoop As Integer)
    Dim i As Integer
    With frmMain
        For i = 0 To iLoop
            cForm.btnFirst(i).Picture = .iPageEnabled.ListImages(1).ExtractIcon
            cForm.btnFirst(i).DisabledPicture = .iPageDisabled.ListImages(1).ExtractIcon
            cForm.btnPrev(i).Picture = .iPageEnabled.ListImages(2).ExtractIcon
            cForm.btnPrev(i).DisabledPicture = .iPageDisabled.ListImages(2).ExtractIcon
            cForm.btnNext(i).Picture = .iPageEnabled.ListImages(3).ExtractIcon
            cForm.btnNext(i).DisabledPicture = .iPageDisabled.ListImages(3).ExtractIcon
            cForm.btnLast(i).Picture = .iPageEnabled.ListImages(4).ExtractIcon
            cForm.btnLast(i).DisabledPicture = .iPageDisabled.ListImages(4).ExtractIcon
            cForm.btnNew(i).Picture = .iPageEnabled.ListImages(5).ExtractIcon
            cForm.btnNew(i).DisabledPicture = .iPageDisabled.ListImages(5).ExtractIcon
            cForm.btnEdited(i).Picture = .iPageEnabled.ListImages(6).ExtractIcon
            cForm.btnEdited(i).DisabledPicture = .iPageDisabled.ListImages(6).ExtractIcon
            cForm.btnAll(i).Picture = .iPageEnabled.ListImages(7).ExtractIcon
            cForm.btnAll(i).DisabledPicture = .iPageDisabled.ListImages(7).ExtractIcon
            
            cForm.btnCN(i).Picture = .i16x16.ListImages(2).ExtractIcon
            cForm.btnES(i).Picture = .i16x16.ListImages(3).ExtractIcon
            cForm.btnS(i).Picture = .i16x16.ListImages(1).ExtractIcon
            cForm.btnD(i).Picture = .i16x16.ListImages(4).ExtractIcon
            cForm.btnR(i).Picture = .i16x16.ListImages(5).ExtractIcon
            cForm.btnP(i).Picture = .i16x16.ListImages(6).ExtractIcon
            cForm.btnC(i).Picture = .i16x16.ListImages(7).ExtractIcon
        Next
    End With
End Sub

Public Sub LoadForm(ByRef srcForm As Form)
    'frmMain.CLOSE_ALL_ACTIVE_FORM
    srcForm.Show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub

Public Function Load_tbHeader(tool_bar As Toolbar, imgList As ImageList, iIcon As Integer)
    'Setting the tool_bar on Imagelist selected or given
    Set tool_bar.ImageList = imgList
    With tool_bar
        'Add Item to Toolbar setting the properties of the Toolbar
        .Buttons.Add , , "Start", 0, iIcon
        '.Buttons.Add , , "", 3
    End With
    
    Set frmMain.tbBR.ImageList = imgList
    With frmMain.tbBR
        'Add Item to Toolbar setting the properties of the Toolbar
        .Buttons.Add , , "Borrow/Return", 0, 20
        '.Buttons.Add , , "", 3
    End With
    
    Set frmMain.tbPenalty.ImageList = imgList
    With frmMain.tbPenalty
        .Buttons.Add , , "Penalty Payment", 0, 19
    End With
End Function

Public Function ToolbarLoader(tBar As Toolbar, sMenu As String, sMenuIcon As String, iLoopMenu As Integer)
    Dim sListMenu As Variant
    Dim sListMenuIcon As Variant
    Dim i As Integer
    sListMenu = Split(sMenu, ",")
    sListMenuIcon = Split(sMenuIcon, ",")
    With tBar
    .Buttons.Clear
    Set .ImageList = frmMain.i16x16
    For i = 0 To iLoopMenu
        .Buttons.Add , , sListMenu(i), , Int(sListMenuIcon(i))
    Next
    End With
End Function


Public Function Activate_tbHeader(tool_bar As Toolbar, strDisabled As String)
    Dim i As Integer
    'this code will check what will be the toolbar button to set disable
    For i = 1 To Len(strDisabled)
        If IsNumeric(Mid(strDisabled, i, 1)) = True Then
            tool_bar.Buttons(Int(Mid(strDisabled, i, 1))).Visible = False
        End If
    Next
End Function

Public Function Show_tbHeader(blShow As Boolean, mdiSelectedFrm As MDIForm)
    If blShow = True Then
        mdiSelectedFrm.picHead.Visible = True
        mdiSelectedFrm.picHeadLn.Visible = True
    Else
        mdiSelectedFrm.picHead.Visible = False
        mdiSelectedFrm.picHeadLn.Visible = False
    End If
End Function

Public Function PIC_LEFT_ON_FOCUS(OBJ_PIC As PictureBox)
    If OBJ_PIC.Width <= 2300 Then
        OBJ_PIC.Width = 2300
        Call frmMain.PIC_RESIZE_LEFT
    End If
End Function

Public Function PIC_LEFT_LOST_FOCUS(OBJ_PIC As PictureBox)
    On Error Resume Next
    If OBJ_PIC.Width >= 0 And OBJ_PIC.Width > 815 Then
        OBJ_PIC.Width = 0
        Call frmMain.PIC_RESIZE_LEFT
    ElseIf OBJ_PIC.Width <= 815 Then
        OBJ_PIC.Width = 0
        Call frmMain.PIC_RESIZE_LEFT
    End If
End Function

Public Function LvPageStat(curren_form As Form, iLv As Integer, sqlStatement As String _
        , iStartPage As Long, iNoPage As Integer, Icon As Integer, sColumns As String _
            , iColumns As Integer, sColumnsWidth As String, sFields As String, sNoRecordMsg As String) As Long
    On Error GoTo errHandler
    
    Dim iRec As Integer
    Dim iNoListPage As Double
    Dim sListColumns As Variant
    Dim sListFields As Variant
    Dim sListColWidth As Variant
    Dim i As Integer
    Dim iLoop As Integer
    Dim mRow As ListItem
    Dim cRec As Long
    Dim iNoPageDec As Integer
    Dim iNoPageVal As Integer
    sListColumns = Split(sColumns, ",")
    sListFields = Split(sFields, ",")
    sListColWidth = Split(sColumnsWidth, ",")
    With curren_form
    Set .lvList(iLv).SmallIcons = frmMain.iLv
    Set .lvList(iLv).Icons = frmMain.iLv
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    .lvList(iLv).ListItems.Clear
    'MsgBox sqlStatement
    adoCon.Open sCon
    adoRes.Open sqlStatement, adoCon, adOpenStatic, adLockOptimistic
        iRec = adoRes.RecordCount
        LvPageStat = iRec
        iNoPageVal = iNoPage
        If iRec > 0 Then
            .lvList(iLv).HideSelection = False
            If Not .lvList(iLv).ColumnHeaders.Count = (iColumns + 2) Then
                .lvList(iLv).ColumnHeaders.Clear
                .lvList(iLv).ColumnHeaders.Add , , , 300
                For i = 0 To iColumns
                    .lvList(iLv).ColumnHeaders.Add , , sListColumns(i), sListColWidth(i)
                Next
            End If
            
            If iRec > iNoPageVal Then
                If iStartPage < 1 Then iStartPage = 1
                
                If (iStartPage + iNoPage) > iRec Then iNoPageVal = iNoPageVal - ((iStartPage + (iNoPageVal - 1)) - iRec)
                
                iNoPageDec = iNoPage
                cRec = 1
                
                Do While Not adoRes.EOF
                    If iStartPage <= cRec And Not iNoPageDec = 0 Then
                        Set mRow = .lvList(iLv).ListItems.Add(, , , , Icon) ' adoRes.Fields(sListFields(0))
                        For i = 0 To iColumns
                            mRow.SubItems(i + 1) = adoRes.Fields(sListFields(i))
                        Next
                        iNoPageDec = iNoPageDec - 1
                    ElseIf iNoPageDec = 0 Then Exit Do
                    End If
                    cRec = cRec + 1
                    adoRes.MoveNext
                Loop
                
                If Not iStartPage = 1 And Not (iStartPage + (iNoPageVal - 1)) = iRec Then
                    .btnFirst(iLv).Enabled = True
                    .btnPrev(iLv).Enabled = True
                    .btnNext(iLv).Enabled = True
                    .btnLast(iLv).Enabled = True
                ElseIf iStartPage = 1 Then
                    .btnFirst(iLv).Enabled = False
                    .btnPrev(iLv).Enabled = False
                    .btnNext(iLv).Enabled = True
                    .btnLast(iLv).Enabled = True
                ElseIf (iStartPage + (iNoPageVal - 1)) = iRec Then
                    .btnFirst(iLv).Enabled = True
                    .btnPrev(iLv).Enabled = True
                    .btnNext(iLv).Enabled = False
                    .btnLast(iLv).Enabled = False
                End If
                
                .lblPageInfo(iLv).Caption = iStartPage & " - " & (iStartPage + (iNoPageVal - 1)) & " of " & iRec
            Else
                Do While Not adoRes.EOF
                    Set mRow = .lvList(iLv).ListItems.Add(, , , , Icon)
                    For i = 0 To iColumns
                        mRow.SubItems(i + 1) = adoRes.Fields(sListFields(i))
                    Next
                    adoRes.MoveNext
                Loop
                .btnFirst(iLv).Enabled = False
                .btnPrev(iLv).Enabled = False
                .btnNext(iLv).Enabled = False
                .btnLast(iLv).Enabled = False
                .lblPageInfo(iLv).Caption = iStartPage & " - " & iRec & " of " & iRec
            End If
            .lblSelected(iLv).Caption = "Selected Record: " & (iStartPage - 1) + .lvList(iLv).SelectedItem.Index
        Else
            .lvList(iLv).HideSelection = True
            If Not .lvList(iLv).ColumnHeaders.Count = 1 Then
                .lvList(iLv).ColumnHeaders.Clear
                .lvList(iLv).ColumnHeaders.Add , , "", 8000
            Else
                .lvList(iLv).ColumnHeaders(1).Text = ""
            End If
            .lvList(iLv).ListItems.Add , , sNoRecordMsg, , 1
            .btnFirst(iLv).Enabled = False
            .btnPrev(iLv).Enabled = False
            .btnNext(iLv).Enabled = False
            .btnLast(iLv).Enabled = False
            .lblPageInfo(iLv).Caption = "0 - 0 of " & iRec
            .lblSelected(iLv).Caption = "Selected Record: None"
        End If
    adoRes.Close
    adoCon.Close
    Set adoRes = Nothing
    Set adoCon = Nothing
    End With
errHandler:
    If err.Number = 94 Then
        Resume Next
    End If
End Function

'Public Function GetMaxID(sSQL As String, sFields) As String
'    Dim mVal As String
'    Dim vMaxId As Variant
'    Set adoCon = New ADODB.Connection
'    Set adoRes = New ADODB.Recordset
'    adoCon.Open conStr
'    adoRes.Open mSql, adoCon, adOpenStatic, adLockOptimistic
'         mVal = adoRes.Fields(mFields)
'    adoRes.Close
'    adoCon.Close
'    Set adoCon = Nothing
'    Set adoRes = Nothing
'    vMaxId = Split(mVal, "-")
'    GET_MAX_ID = vMaxId(1)
'End Function

'This Function is to count the Split item from Expression
Public Function CountSplitItem(Expression As String, Delimeter As String) As Integer
    Dim i As Integer
    CountSplitItem = 0
    For i = 1 To Len(Expression)
        If Mid(Expression, i, 1) = Delimeter Then
            CountSplitItem = CountSplitItem + 1
        End If
    Next
    'CountSplitItem = CountSplitItem + 1
End Function

Public Function INSERT_DATA(strTable As String, strFields As String, strValues As String, strDelimeter As String, bMbox As Boolean)
    Dim sSetValues As String
    Dim SplitValues As Variant
    Dim i As Integer
    'On Error Resume Next
    SplitValues = Split(strValues, strDelimeter)
    For i = 0 To CountSplitItem(strValues, strDelimeter)
        SplitValues(i) = "'" & SplitValues(i) & "'"
    Next
    
    For i = 0 To CountSplitItem(strValues, strDelimeter)
        sSetValues = sSetValues & SplitValues(i)
        If Not i = CountSplitItem(strValues, strDelimeter) Then sSetValues = sSetValues & ","
    Next

    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    'MsgBox "INSERT INTO " & strTable & " (" & strFields & ") VALUES(" & sSetValues & ");"
    adoCon.Open sCon
        adoCon.Execute "INSERT INTO " & strTable & " (" & strFields & ") VALUES(" & sSetValues & ");"
    adoCon.Close
    If bMbox = True Then
        MsgBox "New Record has been save successfully.", vbInformation, "Create New Entry"
    End If
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function UPDATE_DATA(strTable As String, strFields As String, strValues As String, strWhere As String, strDelimeter As String, msgBxStat As Boolean)
    On Error GoTo errHandler
    Dim sSetValues As String
    Dim SplitFields As Variant
    Dim SplitValues As Variant
    Dim sUpdate As String
    Dim i As Integer
    
    'MsgBox strValues
    'MsgBox strFields
    
    SplitValues = Split(strValues, strDelimeter)
    SplitFields = Split(strFields, strDelimeter)
    For i = 0 To CountSplitItem(strValues, strDelimeter)
        SplitValues(i) = "'" & SplitValues(i) & "'"
    Next
    For i = 0 To CountSplitItem(strValues, strDelimeter)
        sSetValues = sSetValues & SplitFields(i) & "=" & SplitValues(i)
        If Not i = CountSplitItem(strValues, strDelimeter) Then sSetValues = sSetValues & ","
    Next
    
    sUpdate = "UPDATE " & strTable & " SET " & sSetValues & " WHERE " & strTable & "." & strWhere & ";"
    'MsgBox sUpdate
    Debug.Print sUpdate
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoCon.Execute sUpdate
    adoCon.Close
    If msgBxStat = True Then
        Call MsgBox("Update Record has been save successfully.", vbInformation, "Update Data")
    End If
    Set adoCon = Nothing
    Set adoRes = Nothing
errHandler:
    If err.Number = -2147467259 Then
        MsgBox err.Number & Chr$(13) & Chr$(13) & " " & err.Description, vbExclamation, "UpdatePointerException"
    End If
End Function

Public Function DELETE_DATA(sTable As String, sField As String, sItem As String)
    Dim sWhere As String, vSplitFields As Variant, vSplitItems As Variant
    Dim i As Integer, sSQL As String
    vSplitFields = Split(sField, ",")
    vSplitItems = Split(sItem, ",")
    For i = 0 To CountSplitItem(sField, ",")
        sWhere = sWhere & sTable & "." & vSplitFields(i) & " like '" & vSplitItems(i) & "'"
        If Not i = CountSplitItem(sField, ",") Then sWhere = sWhere & " And "
    Next
    sSQL = "Delete" & " From " & sTable & " " & _
        "WHERE " & sWhere & ";"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoCon.Execute sSQL
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function FindField(sSQL As String, sFields As String) As String
    On Error Resume Next
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        FindField = adoRes.Fields(sFields)
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function getValueAt(ByVal strUsrid As String) As String
    Dim sSQL As String
    sSQL = "SELECT tbl_users.completename " & _
        "From tbl_users " & _
        "WHERE (((tbl_users.uid) Like '" & strUsrid & "')) " & _
        "GROUP BY tbl_users.completename;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        getValueAt = adoRes.Fields("completename")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Sub HLText(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

Public Function SelectItem(activeForm As frmSelect, sSQL As String, _
                iIcon As Integer, sLblCaption As String, sLblDesc As String, _
                    sXY As String, sTitle As String, sColumns As String, iLoop As Integer, sColWidth As String, _
                        iLvIcon As Integer, sFields As String, sNoRec As String)
    Dim varXY As Variant
    SelectItem = ""
    varXY = Split(sXY, ",")
    With activeForm
        .strSQL = sSQL
        .iIcon = iLvIcon
        .iLoop = iLoop
        .sColumn = sColumns
        .sColWidth = sColWidth
        .sFields = sFields
        .Caption = sTitle
        .Width = varXY(0)
        .Height = varXY(1)
        .sNoRec = sNoRec
        'iStartPage(0) = 1
        'iNoPage(0) = 75
        .imgIcon(0).Picture = frmMain.iListView.ListImages(iIcon).ExtractIcon
        .lblHead(0).Caption = sLblCaption
        .lblDef(0).Caption = sLblDesc
        .Show 1
        SelectItem = .SelectedItem
    End With
End Function

Public Function isNumber(ByVal sKeyAscii As Integer) As Integer
    If sKeyAscii >= 48 And sKeyAscii <= 57 Or sKeyAscii = 8 Then
        isNumber = sKeyAscii
    Else
        isNumber = 0
    End If
End Function

Public Function isNumAndChar(ByVal sKeyAscii) As Integer
    If ((sKeyAscii >= 48 And sKeyAscii <= 57) Or (sKeyAscii >= 65 _
            And sKeyAscii <= 90) Or (sKeyAscii >= 97 And sKeyAscii <= 122) Or (sKeyAscii = 8)) Then
        isNumAndChar = sKeyAscii
    Else
        isNumAndChar = 0
    End If
End Function

Public Function isChar(ByVal sKeyAscii) As Integer
    If (sKeyAscii >= 65 And sKeyAscii <= 90) Or (sKeyAscii >= 97 And sKeyAscii <= 122) Or (sKeyAscii = 8) Then
        isChar = sKeyAscii
    Else
        isChar = 0
    End If
End Function
Public Function isNull(Expression As String) As Boolean
    If Expression = "" Then isNull = True Else isNull = False
End Function

Public Function isCurrency(ByVal sKeyAscii, strCur As String) As Integer
    Dim i As Integer
    Dim intDot As Integer
    intDot = 0
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isCurrency = 0
    Else
        If sKeyAscii = 46 Then
            'You will need this CountSplitItem that i make to run this isCurrency function
            intDot = CountSplitItem(strCur, ".")
            If intDot < 1 Then
                isCurrency = sKeyAscii
            Else
                isCurrency = 0
            End If
        Else
            isCurrency = sKeyAscii
        End If
    End If
End Function

Public Function isRecordExist(strSQL As String) As Boolean
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount < 1 Then
            isRecordExist = False
        Else
            isRecordExist = True
        End If
       ' MsgBox adoRes.RecordCount
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function isCountItem(strSQL As String) As Long
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
        isCountItem = adoRes.RecordCount
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function MatchString(Start As String, sWord As String, bCaseSentive As Boolean) As Boolean
    If bCaseSentive = False Then
        If UCase(Start) = UCase(sWord) Then
            MatchString = True
        Else
            MatchString = False
        End If
    Else
        If Start = sWord Then
            
        End If
    End If
End Function

Public Function FindFieldValue(fSQL As String, fFields As String) As String
    On Error Resume Next
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open fSQL, adoCon, adOpenStatic, adLockOptimistic
        FindFieldValue = adoRes.Fields(fFields)
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function isRecordIfExist(sSQL As String, iLoop As Integer) As Boolean
    Dim i As Integer
    For i = 1 To iLoop
        Set adoCon = New ADODB.Connection
        Set adoRes = New ADODB.Recordset
        adoCon.Open sCon
        adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
            If adoRes.RecordCount < 1 Then
                isRecordIfExist = False
            Else
                isRecordIfExist = True
                Exit For
            End If
        adoRes.Close
        adoCon.Close
        Set adoCon = Nothing
        Set adoRes = Nothing
    Next
End Function

Public Function BarcodeValue() As String
    frmBcodeInputer.Show 1
    BarcodeValue = frmBcodeInputer.sBarcode
End Function

Public Function GenerateID() As String
    GenerateID = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Time, "hhmmss")
End Function

Public Function isBarcodeBookExist(sBarcode As String, bShowMsgBx As Boolean) As Boolean
    Dim sIBBE As String
    sIBBE = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sIBBE, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount < 1 Then
            isBarcodeBookExist = False
        Else
            isBarcodeBookExist = True
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    If bShowMsgBx = True And isBarcodeBookExist = True Then
        MsgBox "Barcode already exist!", vbExclamation, "BarcodePointerException"
    End If
End Function

Public Function isBarcodeMagExist(sBarcode As String, bShowMsgBx As Boolean) As Boolean
    Dim sIBBE As String
    sIBBE = "SELECT tbl_reg_magazines.rm_id " & _
        "From tbl_reg_magazines " & _
        "WHERE (((tbl_reg_magazines.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_magazines.rm_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sIBBE, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount < 1 Then
            isBarcodeMagExist = False
        Else
            isBarcodeMagExist = True
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    If bShowMsgBx = True And isBarcodeMagExist = True Then
        MsgBox "Barcode already exist!", vbExclamation, "BarcodePointerException"
    End If
End Function

Public Function itemCountBorrowed(sB_id As String) As Integer
    Dim sSQL As String
    sSQL = "SELECT tbl_borrow_record.br_id " & _
        "From tbl_borrow_record " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & sB_id & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        itemCountBorrowed = adoRes.RecordCount
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function
'''''''''''''''''''''REFERENCE ONLY''''''''''''''''''
'Public Property Get Fields() As String
'    Fields = m_Fields
'End Property
'
'Set the fields
'Public Property Let Fields(ByVal srcFields As String)
'    m_Fields = srcFields
'End Property

'Public Function ActivateXPTheme(wXP As WindowsXPC)
'    wXP.InitSubClassing
'End Function

Public Function isCanBorrowed(sB_id As String) As Boolean
    Dim sSQL As String, iMaxBorrowed As Integer, iItemBorrowed As Integer
    iMaxBorrowed = GetMaxBorrowed(sB_id)
    iItemBorrowed = itemCountBorrowed(sB_id)
    If iItemBorrowed < iMaxBorrowed Then
        isCanBorrowed = True
    Else
        isCanBorrowed = False
    End If
End Function

Public Function GetMaxBorrowed(sB_id) As Integer
    Dim sSQL As String
    sSQL = "SELECT tbl_borrower_type.maxnoborrow " & _
        "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
        "WHERE (((tbl_borrowers.B_id) Like '" & sB_id & "')) " & _
        "GROUP BY tbl_borrower_type.maxnoborrow;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        GetMaxBorrowed = adoRes.Fields("maxnoborrow")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function ConvertDate(sDate As String) As Date
    ConvertDate = sDate
End Function

Public Function LvSearchItem(fLv As ListView, sItemSearch As String)
    Dim i As Integer
    On Error Resume Next
    For i = 1 To fLv.ListItems.Count
        If fLv.ListItems(i).SubItems(1) = sItemSearch Then
            fLv.ListItems(i).Selected = True
            Exit For
        End If
    Next
End Function

Public Function UPDATE_DATA2(strTable As String, strField As String, strWhere As String, strVal As String)
    Dim strSQL
    strSQL = "UPDATE " & strTable & " SET " & strField & " = '" & strVal & "' " & strWhere & ";"
   ' MsgBox strSQL
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoCon.Execute strSQL
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function
