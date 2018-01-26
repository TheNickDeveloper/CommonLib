VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonFunctionHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================

'Exception Catch

'=============================================

Public Function ExceptionError()
    If Err.Number Then
         MsgBox Err.Description, vbOKOnly, "Error" & " - " & Err.Source
    End If
End Function

'=============================================

'Data manipulating function

'=============================================

'get the value between strings
Public Function GetValueBetweenString(ByVal Value As String, ByVal Mark_1 As String, ByVal Mark_2 As String) As Variant
    Dim openPos As Integer, closePos As Integer, Rseult As String
    
    If Len(Value) Then
        openPos = InStr(Value, Mark_1)
        closePos = InStr(Value, Mark_2)
        GetValueBetweenString = Mid(Value, openPos + 1, closePos - openPos - 1)
    End If
End Function

'get Transfer To Date
Public Function GetTransToDate(ByVal lngTargetDate As Variant, Optional ByVal formatText As String = "") As Variant
    Dim formatDate As Variant
    
    If Len(lngTargetDate) = 8 And IsNumeric(lngTargetDate) Then
        formatDate = DateSerial(Left(lngTargetDate, 4), Mid(lngTargetDate, 5, 2), Right(lngTargetDate, 2))
        GetTransToDate = formatDate
    Else
        GetTransToDate = lngTargetDate
    End If
End Function

'put value from array to dic
Public Function SetDicFromArray(ByVal arrTarget As Variant, ByVal inputKeyIndex As String, ByVal inputValueIndex As String, Optional ByVal delimiteChar As String = ",") As Object
    Dim i As Long, j As Long, Key As Variant, Value As Variant, arrTemp As Variant
    Dim TargetDic As New Dictionary, arrKeyIndex() As String, arrResultIndex() As String
    
    arrKeyIndex = Split(inputKeyIndex, ",")
    arrResultIndex = Split(inputValueIndex, ",")
    
    For i = 1 To UBound(arrTarget)
        For j = LBound(arrKeyIndex) To UBound(arrKeyIndex)
            Key = Key & arrTarget(i, arrKeyIndex(j))
        Next
        
        For j = LBound(arrResultIndex) To UBound(arrResultIndex)
            If j = 0 Then
                Value = arrTarget(i, arrResultIndex(j))
            Else
                Value = Value & delimiteChar & arrTarget(i, arrResultIndex(j))
            End If
        Next
    
        If Not TargetDic.Exists(Key) Then
            TargetDic.Add Key, Value
        End If
        
        Key = ""
        Value = ""
    Next
    
    Set SetDicFromArray = TargetDic
End Function

'=============================================

'Excel Workbook & Worksheet function

'=============================================

'Column Count
Public Function GetUsedColNum(ByVal wsName As Worksheet, Optional ByVal lngStartRow As Long = 1) As Long
    GetUsedColNum = wsName.Cells(lngStartRow, Columns.Count).End(xlToLeft).Column
End Function

'Row Count
Public Function GetUsedRowNum(ByVal wsName As Worksheet, Optional ByVal lngStartCol As Long = 1) As Long
    GetUsedRowNum = wsName.Cells(Rows.Count, lngStartCol).End(xlUp).Row
End Function

'from worksheet to array
Public Function ConvertTargetFromRangeToArray(ByVal wsTarget As Worksheet, ByVal startRow As Integer, startCol As Integer) As Variant
    Dim usedRow As Long, usedCol As Long, arrTemp As Variant

    usedRow = GetUsedRowNum(wsTarget, startCol)
    usedCol = GetUsedColNum(wsTarget, startRow)
    With wsTarget
        arrTemp = .Range(.Cells(startRow, startCol), .Cells(usedRow, usedCol))
    End With
    ConvertTargetFromRangeToArray = arrTemp
End Function

'save As Worksheet Contents Into New Workbook
Public Sub SaveAsWorksheetContentsIntoNewWorkbook(ByVal wsTarget As Worksheet, ByVal startRow As Integer, ByVal startCol As Integer, ByVal pathOutput, ByVal wsTargetName As String)
    Dim arrTemp As Variant
    
    arrTemp = ConvertTargetFromRangeToArray(wsTarget, startRow, startCol)
    Workbooks.Add
    
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .Worksheets(1).[A1].Resize(UBound(arrTemp), UBound(arrTemp, 2)) = arrTemp
        .BuiltinDocumentProperties("Comments").Value = "INTERNAL"
        .SaveAs FileName:=pathOutput & "\" & wsTargetName & ".xlsx"
        .Close
    End With
    Application.DisplayAlerts = True
End Sub

'=============================================

'UI related function

'=============================================

Public Function GetFileOrFolderPath(ByVal TargetCategory As String)
    Dim CorrespondingPerameter As Variant
    
    If TargetCategory = "FilePath" Then
        CorrespondingPerameter = msoFileDialogOpen
    Else
        CorrespondingPerameter = msoFileDialogFolderPicker
    End If
    
    With Application.FileDialog(CorrespondingPerameter)
        .AllowMultiSelect = True
        If .Show = True Then
            GetFileOrFolderPath = .SelectedItems(1)
        Else
            MsgBox "Action has been cancelled."
        End If
    End With
End Function

'=============================================

'ADO Helper Assistant

'=============================================

Public Function GetExcelRecordSetArray(ByRef ADO As Object, ByVal wsSourcePath As String, ByVal strQuery, Optional ByVal includeHeader As Boolean = True) As Variant
    With ADO
        .ConnectExcel wsSourcePath
        
        If Len(.GetErrMessage) = 0 Then
            .OpenRecordSet strQuery
            GetRecordSetArray = .GetAllRowsData(includeHeader)
        End If
    End With
End Function

Public Function GetTxtRecordSetArray(ByRef ADO As Object, ByVal sourcePath As String, ByVal strQuery, Optional ByVal includeHeader As Boolean = True) As Variant
    Dim ParentPath As String
    
    ParentPath = Left$(sourcePath, InStrRev(sourcePath, "\"))
    
    With ADO
        .connectText ParentPath, includeHeader
        
        If Len(.GetErrMessage) = 0 Then
            .OpenRecordSet strQuery
            GetTxtRecordSetArray = .GetAllRowsData(includeHeader)
        End If
    End With
End Function



