VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AS400Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputBuf As GdiplusStartupInput, ByVal outputBuf As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, ByRef bitmap As Long) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpszCLSID As Long, ByRef pCLSID As GUID) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    Type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Public Enum CaptureKeyStyle
    EqualKey
    containkey
End Enum

Private Const SRCCOPY = 13369376
Private Const ERR_WAIT = "Wait for timeout"
Private Const ERR_FIELDPROTECTED = "Field was protected, can not send any contents to it."
Private Const ERR_FIELDNOTFUND = "Field not found."
Private Const ERRSESSION = "Session name doesn't exist."

Private autECLSession As Object
Private autECLConnList  As Object

Private nameOfSession As String
Private errorCapture As String
Private intervalSleepTime As Long
Private maxSleepTime As Long
Private maxCycleNum As Long

Private maxRowNum As Long
Private maxColNum As Long

Private winHandle As Long
Private picQuality As Integer
Private picSaveFolder As String

Private isReady As Boolean
Private IsRightAlign As Boolean
    
'--------------------------------------------------------------
'**************************************************************
'
' Class Initialize & Terminate
'
'**************************************************************
'--------------------------------------------------------------

Private Sub Class_Initialize()
    'create PComm object
    Set autECLSession = CreateObject("PCOMM.autECLSession")
    Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    
    'set property default value
    intervalSleepTime = 20
    maxSleepTime = 60000
    nameOfSession = ""
    
    maxRowNum = 0
    maxColNum = 0
    winHandle = 0
    picQuality = 50
    errorCapture = ""
    isReady = False
    IsRightAlign = False
    picSaveFolder = Environ("Temp") & "\"
End Sub

Private Sub Class_Terminate()
    Set autECLSession = Nothing
    Set autECLConnList = Nothing
End Sub

'--------------------------------------------------------------
'**************************************************************
'
' Class's Property
'
'**************************************************************
'--------------------------------------------------------------

Public Property Let ContentsRightAlign(ByVal val As Boolean)
    IsRightAlign = val
End Property

Public Property Let MaxWaitTime(ByVal val As Long)
    If val > 0 Then
        maxSleepTime = val
        maxCycleNum = maxSleepTime \ intervalSleepTime
    End If
End Property

Public Property Get MaxWaitTime() As Long
    MaxWaitTime = maxSleepTime
End Property

Public Property Let IntervalDetectTime(ByVal val As Long)
    If val > 0 And val < maxSleepTime Then
        intervalSleepTime = val
        maxCycleNum = maxSleepTime \ intervalSleepTime
    End If
End Property

Public Property Get IntervalDetectTime() As Long
    IntervalDetectTime = intervalSleepTime
End Property

Public Property Let SnapshotSaveFolder(ByVal val As String)
    If Len(Dir(val, vbDirectory)) Then
        If Right(val, 1) <> "\" Then
            val = val & "\"
        End If
        picSaveFolder = val
    End If
End Property

Public Property Get SnapshotSaveFolder() As String
    SnapshotSaveFolder = picSaveFolder
End Property

Public Property Let SnapshotQuality(ByVal val As Integer)
    If val > 0 And val <= 100 Then
        picQuality = val
    End If
End Property

Public Property Get MaxRows() As Long
    MaxRows = maxRowNum
End Property

Public Property Get MaxCols() As Long
    MaxCols = maxColNum
End Property

Public Property Get EmulatorWindowHandle() As Long
    EmulatorWindowHandle = winHandle
End Property

Public Property Get getErrorMessage() As String
    getErrorMessage = errorCapture
End Property

Public Property Get isConnected() As Boolean
    isConnected = isReady
End Property

'--------------------------------------------------------------
'**************************************************************
'
' Class Public Function & Sub
'
'**************************************************************
'--------------------------------------------------------------

Public Sub SleepWait(ByVal sleepTime As Long)
    Sleep sleepTime
End Sub

Public Sub SetCursor(ByVal rowNum As Integer, ByVal colNum As Integer)
    If isReady Then
        If rowNum > 0 And colNum > 0 Then
            If rowNum <= maxRowNum And colNum <= maxColNum Then
                autECLSession.autECLPS.SetCursorPos rowNum, colNum
            End If
        End If
    End If
End Sub

Public Sub SendKeys(ByVal strKeys As String, Optional ByVal sleepTime As Long = 20)
    If isReady Then
        If Left(strKeys, 1) = "[" And Right(strKeys, 1) = "]" Then
            Call SendKeyToScreen(strKeys, sleepTime)
        End If
    End If
End Sub

Public Function ConnectSession(ByVal SessionName As String) As Boolean
    Dim cycleCount As Integer
    
    nameOfSession = Trim(SessionName)
    isReady = False
    errorCapture = ""
    
    If IsSessionExists(nameOfSession) Then
        maxCycleNum = maxSleepTime \ intervalSleepTime

        If Len(SessionName) Then
            autECLSession.SetConnectionByName nameOfSession
            
            cycleCount = 1
            Do
                Sleep intervalSleepTime
                DoEvents
                If autECLSession.autECLPS.Ready = False Then
                    cycleCount = cycleCount + 1
                End If
            Loop Until autECLSession.autECLPS.Ready Or cycleCount > maxCycleNum
            
            isReady = autECLSession.autECLPS.Ready
        End If
        
        If isReady Then
            maxRowNum = autECLSession.autECLPS.NumRows
            maxColNum = autECLSession.autECLPS.NumCols
            winHandle = GetHandle
        Else
            nameOfSession = ""
            errorCapture = ERR_WAIT
        End If
    Else
        errorCapture = ERRSESSION
    End If
    
    ConnectSession = isReady
End Function

Public Function BackToSpecifyScreen(ByVal fieldKey As String, Optional ByVal rowNum As Integer = 0, Optional ByVal captureStyle As CaptureKeyStyle = containkey, Optional ByVal waitScreenRender As Long = 20) As Boolean
    Dim isBackTo As Boolean, isTopScreen As Boolean
    Dim iStep As Long

    isBackTo = False
    isTopScreen = False
    If isReady Then
        If Len(Trim(GetContents(MaxRows, 1, MaxCols))) Then
            SendKeys "[RESET]"
        End If

        Do
            If CheckScreenInputInhibited(iStep, waitScreenRender) Then
                If DetctKeyField(fieldKey, rowNum, captureStyle) Then
                    isBackTo = True
                Else
                    For iStep = MaxRows To 1 Step -1
                        If DetctKeyField("F1-Exit", iStep, EqualKey) Then
                            Exit For
                        End If
                    Next

                    If iStep = 0 Then
                        isTopScreen = True
                    Else
                        SendKeys "[PF1]"
                    End If
                End If
            End If
        Loop Until isTopScreen Or isBackTo
    End If

    BackToSpecifyScreen = isBackTo
End Function

Public Function SendContent(ByVal valContents As String, ByVal rowNum As Integer, ByVal colNum As Integer, Optional ByVal IsSendEnterKey As Boolean = True) As Boolean
    Dim isSuccess As Boolean
    
    isSuccess = False
    Err.Clear
    On Error GoTo ErrHandle
    
    If isReady Then
        If rowNum > 0 And colNum > 0 Then
            If IsFieldExists(rowNum, colNum) Then
                autECLSession.autECLPS.SetCursorPos rowNum, colNum
                If IsRightAlign Then
                    SendKeyToScreen "[eraseeof]"
                    SendKeyToScreen valContents
                    SendKeyToScreen "[field+]"
                Else
                    SendKeyToScreen "[eraseeof]"
                    autECLSession.autECLPS.SetText valContents, rowNum, colNum
                End If
                
                If IsSendEnterKey Then
                    SendKeyToScreen "[Enter]"
                End If
                isSuccess = True
            End If
        End If
        
ErrHandle:
        If Err.Number Then
            errorCapture = Err.Description
            isSuccess = False
        End If
    End If
    
    SendContent = isSuccess
End Function

Public Function SendMultipleContent(ByRef arrayContents As Variant, Optional ByVal IsSendEnterKey As Boolean = True) As Boolean
    Dim iStep As Integer, isSuccess As Boolean

    For iStep = LBound(arrayContents) To UBound(arrayContents) Step 3
        isSuccess = SendContent(arrayContents(iStep), arrayContents(iStep + 1), arrayContents(iStep + 2), False)
        If isSuccess = False Then
            Exit For
        End If
    Next
    
    If isSuccess And IsSendEnterKey Then
        SendKeyToScreen "[Enter]"
    End If
    
    SendMultipleContent = isSuccess
End Function

Public Sub WaitSpecifyText(ByVal rowNum As Integer, ByVal colNum As Integer)
    
End Sub

Public Function WaitScreenShow(ByVal fieldKey As String, Optional ByVal rowNum As Integer = 0, Optional ByVal captureStyle As CaptureKeyStyle = containkey) As Boolean
    Dim isShow As Boolean
    Dim waitTime As Long
    
    If isReady Then
        isShow = False
        waitTime = 0
        Do
            Sleep 50
            waitTime = waitTime + 50
            isShow = DetctKeyField(fieldKey, rowNum, captureStyle)
        Loop Until isShow Or waitTime > maxSleepTime
    End If
    
    WaitScreenShow = isShow
End Function

Public Function DetectScreenCompleted(ByVal fieldKey As String, Optional ByVal rowNum As Integer = 0, Optional ByVal captureStyle As CaptureKeyStyle = containkey, Optional ByVal waitScreenRender As Long = 20) As Boolean
    Dim cycleCount As Long, IsComplete As Boolean

    errorCapture = ""
    If isReady Then
        IsComplete = CheckScreenInputInhibited(cycleCount, waitScreenRender)
        If IsComplete Then
            If DetctKeyField(fieldKey, rowNum, captureStyle) = False Then
                IsComplete = False
                errorCapture = Trim(GetContents(maxRowNum, 1, maxColNum))
            End If
        Else
            If cycleCount > maxCycleNum Then
                errorCapture = ERR_WAIT
            End If
        End If
    End If
    
    DetectScreenCompleted = IsComplete
End Function

'return a variant array: col1=series number, col2=start row, col3=start column, col4=field length, col5=field type, col6=field contents
Public Function GetScreenFieldsInfo() As Variant()
    Dim iCount As Integer, arrResult() As Variant
    Dim iStep As Integer
    
    If isReady Then
        With autECLSession.autECLPS
            .autECLFieldList.Refresh
            
            iCount = .autECLFieldList.Count
            If iCount Then
                ReDim arrResult(1 To iCount, 1 To 6)
                
                For iStep = 1 To iCount
                    'get fields information: Row, Column, Length, Field attribute
                    arrResult(iStep, 1) = iStep
                    arrResult(iStep, 2) = .autECLFieldList(iStep).startRow
                    arrResult(iStep, 3) = .autECLFieldList(iStep).startCol
                    arrResult(iStep, 4) = .autECLFieldList(iStep).Length
                    arrResult(iStep, 5) = IIf(.autECLFieldList(iStep).Protected, "Protected", "Modified")
                    arrResult(iStep, 6) = GetContents(arrResult(iStep, 2), arrResult(iStep, 3), arrResult(iStep, 4))
                Next
                
                GetScreenFieldsInfo = arrResult
            End If
        End With
    End If
End Function

Public Function GetContents(ByVal rowNum As Integer, ByVal colNum As Integer, ByVal contentLength As Integer) As String
    Dim strTemp As String, detectNum As Long
    
    strTemp = ""
    If isReady Then
        If CheckScreenInputInhibited(detectNum, intervalSleepTime) Then
            strTemp = autECLSession.autECLPS.GetText(rowNum, colNum, contentLength)
        End If
    End If
    
    GetContents = strTemp
End Function

Public Function SnapScreen(Optional ByVal saveName As String = "", Optional ByVal waitScreenRender As Long = 20) As String
    Dim savePath As String, detectNum As Long
    
    Err.Clear
    On Error GoTo ErrHandle
    savePath = saveName
    
    If isReady Then
        If CheckScreenInputInhibited(detectNum, waitScreenRender) Then
            Call SnapShort(savePath)
ErrHandle:
            If Err.Number Then
                savePath = ""
                errorCapture = Err.Description
            End If
        End If
    End If
    
    SnapScreen = savePath
End Function

Public Function ScreenInsertToWorkSheet(ByVal rangeTarget As Range, Optional ByVal wsBack As Worksheet = Nothing) As Boolean
    Dim picPath As String, isSuccess As Boolean
    
    Err.Clear
    On Error GoTo ErrHandle
    isSuccess = False
    
    If isReady Then
        picPath = SnapScreen
        If Len(picPath) Then
            With rangeTarget.Parent
                .Select
                rangeTarget.Select
                .Pictures.Insert picPath
            End With
            
            If Not wsBack Is Nothing Then
                wsBack.Select
            End If
            
            isSuccess = True
        End If
    End If
    
ErrHandle:
    If Len(picPath) Then
        Kill picPath
    End If
    
    If Err.Number Then
        errorCapture = Err.Description
        isSuccess = False
    End If
    
    ScreenInsertToWorkSheet = isSuccess
End Function

'--------------------------------------------------------------
'**************************************************************
'
' Internal Function
'
'**************************************************************
'--------------------------------------------------------------

Private Function CheckScreenInputInhibited(ByRef detctCount As Long, ByVal sleepTime As Long) As Boolean
    Dim flagNum As Integer, IsComplete As Boolean
    
    detctCount = 1
    IsComplete = False
    
    Do
        flagNum = autECLSession.autECLOIA.InputInhibited
        '0 = Not Inhibited
        '5 = Other Inhibit
        If flagNum = 0 Or flagNum = 5 Then
            IsComplete = True
        Else
            Sleep intervalSleepTime
            DoEvents
        End If

        detctCount = detctCount + 1
    Loop Until detctCount > maxCycleNum Or IsComplete
    
    If IsComplete Then
        If sleepTime > 0 Then
            Sleep sleepTime
        End If
    End If
    
    CheckScreenInputInhibited = IsComplete
End Function

Private Function IsFieldExists(ByVal rowNum As Integer, ByVal colNum As Integer) As Boolean
    Dim iStep As Long, iCount As Long
    Dim isFound As Boolean, IsModified As Boolean
    
    isFound = False
    IsModified = False
    
    With autECLSession.autECLPS
        .autECLFieldList.Refresh
        iCount = .autECLFieldList.Count
    
        If iCount Then
            For iStep = 1 To iCount
                With .autECLFieldList(iStep)
                    'detect fields row number and column number were correct and field's attribute is modified
                    If (.startRow = rowNum) And (.startCol = colNum) Then
                        isFound = True
                        If .Protected = False Then
                            IsModified = True
                            Exit For
                        End If
                    ElseIf .startRow > rowNum Then
                        Exit For
                    End If
                End With
            Next
        End If
    End With
    
    If isFound Then
        If IsModified = False Then
            errorCapture = ERR_FIELDPROTECTED & " RowNum: " & rowNum & " ColNum: " & colNum
        End If
    Else
        errorCapture = ERR_FIELDNOTFUND & " RowNum: " & rowNum & " ColNum: " & colNum
    End If
    
    IsFieldExists = IsModified
End Function

Private Function IsSessionExists(ByVal SessionName As String) As Boolean
    Dim iStep As Integer, iCount As Integer, IsExists As Boolean
    
    IsExists = False
    
    autECLConnList.Refresh
    iCount = autECLConnList.Count
    
    For iStep = 1 To iCount
        If autECLConnList(iStep).Name = SessionName Then
            IsExists = True
            Exit For
        End If
    Next
    
    IsSessionExists = IsExists
End Function

Private Sub SendKeyToScreen(ByVal strKeys As String, Optional ByVal sleepTime As Long = 0)
    autECLSession.autECLPS.SendKeys strKeys
    If sleepTime > 0 Then
        Sleep sleepTime
    End If
End Sub

Private Function GetHandle() As Long
    Dim autECLWinObj As Object
    
    Set autECLWinObj = CreateObject("PCOMM.autECLWinMetrics")

    autECLWinObj.SetConnectionByName nameOfSession
    GetHandle = FindWindow(vbNullString, autECLWinObj.WindowTitle)
    
    Set autECLWinObj = Nothing
End Function

Private Function DetctKeyField(ByVal keyWord As String, ByVal rowNum As Integer, ByVal captureStyle As CaptureKeyStyle) As Boolean
    Dim iCount As Integer, iStep As Integer, IsCapture As Boolean
    Dim fieldContents As String
    
    IsCapture = False
    
    If Len(keyWord) And rowNum <= maxRowNum Then
        keyWord = UCase(keyWord)
        With autECLSession.autECLPS
            .autECLFieldList.Refresh
            
            iCount = .autECLFieldList.Count
            If rowNum > 0 Then
                For iStep = 1 To iCount
                    If .autECLFieldList(iStep).startRow = rowNum Then
                        fieldContents = UCase(GetContents(.autECLFieldList(iStep).startRow, .autECLFieldList(iStep).startCol, .autECLFieldList(iStep).Length))
                        If captureStyle = EqualKey Then
                            IsCapture = keyWord = fieldContents
                        ElseIf captureStyle = containkey Then
                            IsCapture = InStr(fieldContents, keyWord)
                        End If
                        
                        If IsCapture Then
                            Exit For
                        End If
                    ElseIf .autECLFieldList(iStep).startRow > rowNum Then
                        Exit For
                    End If
                Next
            Else
                For iStep = 1 To iCount
                    fieldContents = UCase(GetContents(.autECLFieldList(iStep).startRow, .autECLFieldList(iStep).startCol, .autECLFieldList(iStep).Length))
                    If captureStyle = EqualKey Then
                        IsCapture = keyWord = fieldContents
                    ElseIf captureStyle = containkey Then
                        IsCapture = InStr(fieldContents, keyWord)
                    End If
                    
                    If IsCapture Then
                        Exit For
                    End If
                Next
            End If
        End With
    End If
    
    DetctKeyField = IsCapture
End Function

Private Sub SnapShort(ByRef savePath As String)
    Dim actDC As Long, comDC As Long
    Dim dcWidth As Long, dcHeight As Long
    Dim scrMap As Long, selectObj As Long
    Dim actRECT As RECT
    
    actDC = GetWindowDC(winHandle)
    comDC = CreateCompatibleDC(actDC)
    
    GetWindowRect winHandle, actRECT

    dcHeight = actRECT.Bottom - actRECT.Top
    dcWidth = actRECT.Right - actRECT.Left
    
    scrMap = CreateCompatibleBitmap(actDC, dcWidth, dcHeight)
    selectObj = SelectObject(comDC, scrMap)
    BitBlt comDC, 0, 0, dcWidth, dcHeight, actDC, 0, 0, SRCCOPY

    If Len(savePath) = 0 Then
        savePath = Timer * 100
    End If
    
    savePath = picSaveFolder & savePath & ".jpg"
    Call SavePic(scrMap, savePath)
    
    DeleteDC comDC
    ReleaseDC winHandle, actDC
End Sub

Private Sub SavePic(ByVal lngMap As Long, ByVal sPath As String)
    Dim GSI As GdiplusStartupInput
    Dim GDI As Long, retResult As Long, nMap As Long
    Dim picEncoder As GUID, eParams As EncoderParameters
    
    GSI.GdiplusVersion = 1
    retResult = GdiplusStartup(GDI, GSI, 0)
    If retResult = 0 Then
        retResult = GdipCreateBitmapFromHBITMAP(lngMap, 0, nMap)
        If retResult = 0 Then
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), picEncoder
            
            eParams.Count = 1
            With eParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .Type = 4
                .Value = VarPtr(picQuality)
            End With
            
            retResult = GdipSaveImageToFile(nMap, StrPtr(sPath), picEncoder, eParams)
            
            GdipDisposeImage nMap
        End If
    End If
End Sub
