VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Enum dmVandas
    vbline = 1
    vbColumn = vbline + 1
    vbBoth = vbColumn + 1
    vbInLine = vbBoth + 1
    vbInColumn = vbInLine + 1
    vbDescendingOrder = vbInColumn + 1
    vbAscendingOrder = vbDescendingOrder + 1
End Enum

Private Const MAXLONG As Long = (2 ^ 31) - 1
Private vmData As Variant, bmAsInline As Boolean, lmDimention As Long, vmColumnNames As Variant, vmLineNames As Variant, _
    lmDimentions As Long, lmDesiredDimention As Long
Private Const C_Column As String = "column", C_Row As String = "row", C_StringType As String = "STRING", _
    C_IntegerType As String = "INTEGER", C_LongType As String = "LONG", C_ArrayType As String = "VARIANT()", _
    CSV_SIGN As String = ";", C_Dot As String = ".", C_Comma As String = ","

Public Property Let data(vData As Variant)
    vmData = adjustData(vData)
    columnHeaders = vbEmpty
    rowHeaders = vbEmpty
End Property

Public Property Get data() As Variant
    data = vmData
End Property

Public Property Get column(Optional ByVal InLine As Integer = -1, Optional ByVal columns As Variant) As Object
    dimention = dmVandas.vbColumn
    asInline = IIf(InLine = dmVandas.vbInLine, True, False)
    Set column = Me
End Property

Public Property Get row(Optional ByVal InLine As Integer = -1, Optional ByVal rows As Variant) As Object
    dimention = dmVandas.vbline
    asInline = IIf(InLine = dmVandas.vbInLine, True, False)
    Set row = Me
End Property

Private Property Let asInline(bAsInLine As Boolean)
    bmAsInline = bAsInLine
End Property
'
Private Property Get asInline() As Boolean
    asInline = bmAsInline
End Property

Private Property Let dimention(lDimention As Long)
    lmDesiredDimention = lDimention
End Property

Private Property Get dimention() As Long
    dimention = lmDesiredDimention
End Property

Private Property Let classDimention(lDimention As Long)
    lmDimention = lDimention
End Property

Private Property Get classDimention() As Long
    classDimention = lmDimention
End Property

Private Property Let inputDimention(lDimentions As Long)
    lmDimentions = lDimentions
End Property

Private Property Get inputDimention() As Long
    inputDimention = lmDimentions
End Property

Private Property Get columnHeaders() As Variant
    columnHeaders = vmColumnNames
End Property

Private Property Get rowHeaders() As Variant
    rowHeaders = vmLineNames
End Property

Private Property Let columnHeaders(ByVal vHeaders As Variant)
    vmColumnNames = vHeaders
End Property

Private Property Let rowHeaders(ByVal vHeaders As Variant)
    vmLineNames = vHeaders
End Property


Public Property Get length() As Long
    Dim lDimention As Long

    lDimention = classDimention
    If lDimention = dmVandas.vbColumn Or lDimention = dmVandas.vbline Then
        length = lengthDimention(lDimention)
    Else
        length = lengthMatrix
    End If

    classDimention = dmVandas.vbBoth
End Property

Private Function lengthDimention(ByVal iDimention As Integer) As Long
    Dim lResult As Long, lStart As Long
    
    If getDimentions < iDimention Then lengthDimention = 0: Exit Function
    
    lStart = LBound(data, iDimention)
    lResult = UBound(data, iDimention) - lStart
    
    lengthDimention = IIf(lResult > 0, lResult + 1, 0) ' +1 to include the last element
End Function

Private Function getColumns(ByVal vColumns As Variant) As Variant
    Dim vResult As Variant, vColumn As Variant
    Dim lRowCount As Long, lColumnCount As Long, lColumn As Long
    Dim lDimention As Long, lRows As Long, lColumns As Long, lDelta As Long, bChanged As Boolean
    
    classDimention = dmVandas.vbColumn: lDimention = length
    classDimention = dmVandas.vbline: lRows = length
    lDelta = UBound(vColumns) - LBound(vColumns)
    lDelta = lDelta + 1
    lColumns = UBound(vColumns)
    ReDim vResult(1 To lRows, 1 To lDelta)
    For lRowCount = 1 To lRows
        lColumnCount = 1
        For Each vColumn In vColumns
            If Not IsNumeric(vColumn) Then getColumns = Array(): Exit Function
            lColumn = CLng(vColumn)
            If lColumn > 0 And lColumn <= lDimention Then   'getColumns = Array(): Exit Function
                vResult(lRowCount, lColumnCount) = vmData(lRowCount, lColumn)
                bChanged = True
            End If
            lColumnCount = lColumnCount + 1
        Next vColumn
    Next lRowCount

    getColumns = IIf(bChanged, vResult, Array())
End Function

Private Function getRows(ByVal vRows As Variant) As Variant
    Dim vResult As Variant, vRow As Variant
    Dim lRowCount As Long, lColumnCount As Long, lRow As Long
    Dim lDimention As Long, lRows As Long, lColumns As Long, lDelta As Long, bChanged As Boolean
 
    classDimention = dmVandas.vbline: lDimention = length
    classDimention = dmVandas.vbColumn: lColumns = length
    lDelta = UBound(vRows) - LBound(vRows)
    lDelta = lDelta + 1
    lRows = UBound(vRows)
    ReDim vResult(1 To lDelta, 1 To lColumns)
    For lColumnCount = 1 To lColumns
        lRowCount = 1
        For Each vRow In vRows
            If Not IsNumeric(vRow) Then getRows = Array(): Exit Function
            lRow = CLng(vRow)
            If lRow > 0 And lRow <= lDimention Then 'getRows = Array(): Exit Function
                vResult(lRowCount, lColumnCount) = vmData(lRow, lColumnCount)
                bChanged = True
            End If
            lRowCount = lRowCount + 1
        Next vRow
    Next lColumnCount

    getRows = IIf(bChanged, vResult, Array())
End Function

Public Function columns(ParamArray args() As Variant) As Variant
    Dim vitem As Variant, vList As Variant
    Dim sType As String
    
    For Each vitem In args
        sType = UCase(TypeName(vitem))
        If sType = C_ArrayType Then
            columns = getIntervalColumns(vitem): Exit Function
        End If
        
        If sType = C_StringType Then
            vList = args: columns = getNamedColumns(vList): Exit Function
        End If
        
        If sType = C_IntegerType Or sType = C_LongType Then
            vList = args: columns = getColumns(vList): Exit Function
        End If
    Next vitem
    columnHeaders = vbEmpty
    columns = Array()
End Function

Public Function rows(ParamArray args() As Variant) As Variant
    Dim vitem As Variant, vList As Variant
    Dim sType As String
    
    For Each vitem In args
        sType = UCase(TypeName(vitem))
        If sType = C_ArrayType Then
            rows = getIntervalRows(vitem): Exit Function
        End If
        
        If sType = C_StringType Then
            vList = args: rows = getNamedRows(vList): Exit Function
        End If
        
        If sType = C_IntegerType Or sType = C_LongType Then
            vList = args: rows = getRows(vList): Exit Function
        End If
    Next vitem
    rowHeaders = vbEmpty
    rows = Array()
End Function

Private Function getNamedColumns(ByRef vColumnNames As Variant) As Variant
    Dim vColumns As Variant, vDataNames As Variant, vitem As Variant
    Dim lPosition As Long
    Dim sList As String, sSign

    vDataNames = getColumnNames
    sList = vbNullString
    sSign = vbNullString
    For Each vitem In vColumnNames
        lPosition = getArrayPosition(vitem, vDataNames)
        If lPosition > 0 Then sList = sList & sSign & CStr(lPosition): sSign = CSV_SIGN
    Next vitem
    
    vColumns = IIf(InStr(sList, CSV_SIGN) > 0, Split(sList, CSV_SIGN), Array(sList))
    
    getNamedColumns = getColumns(vColumns)
End Function

Private Function getColumnNames() As Variant
    Dim vResult As Variant
    Dim lColumns As Long
    Const C_HeaderRow As Integer = 1
    
    If isInitializedArray(vmColumnNames) Then _
        getColumnNames = vmColumnNames: Exit Function
    
    classDimention = dmVandas.vbColumn: lColumns = length
    ReDim vResult(1 To lColumns) As Variant
    While lColumns > 0
        vResult(lColumns) = vmData(C_HeaderRow, lColumns)
        lColumns = lColumns - 1
    Wend
    
    vmColumnNames = vResult
    getColumnNames = vResult
End Function

Private Function getNamedRows(ByRef vRowsNames As Variant) As Variant
    Dim vRows As Variant, vDataNames As Variant, vitem As Variant
    Dim lPosition As Long
    Dim sList As String, sSign

    vDataNames = getLineNames
    sList = vbNullString
    sSign = vbNullString
    For Each vitem In vRowsNames
        lPosition = getArrayPosition(vitem, vDataNames)
        If lPosition > 0 Then sList = sList & sSign & CStr(lPosition): sSign = CSV_SIGN
    Next vitem
    
    vRows = IIf(InStr(sList, CSV_SIGN) > 0, Split(sList, CSV_SIGN), Array(sList))
    
    getNamedRows = getRows(vRows)
End Function

Private Function getLineNames() As Variant
    Dim vResult As Variant
    Dim lRows As Long
    Const C_HeaderColumn As Integer = 1
    
    If isInitializedArray(vmLineNames) Then _
        getLineNames = vmLineNames: Exit Function
    
    
    classDimention = dmVandas.vbline: lRows = length
    ReDim vResult(1 To lRows) As Variant
    While lRows > 0
        vResult(lRows) = vmData(lRows, C_HeaderColumn)
        lRows = lRows - 1
    Wend
    
    vmLineNames = vResult
    getLineNames = vResult
End Function

Private Function getArrayPosition(ByRef vitem As Variant, ByVal vList As Variant) As Long
    Dim vListItem As Variant
    Dim lResult As Long
    
    lResult = 1
    For Each vListItem In vList
        If vitem = vListItem Then getArrayPosition = lResult: Exit Function
        lResult = lResult + 1
    Next vListItem
    
    getArrayPosition = 0
End Function

Private Function getIntervalColumns(ByVal vInterval As Variant) As Variant
    getIntervalColumns = getInterval(vInterval, dmVandas.vbColumn)
End Function

Private Function getIntervalRows(ByVal vInterval As Variant) As Variant
    getIntervalRows = getInterval(vInterval, dmVandas.vbline)
End Function

Private Function getInterval(ByVal vInterval As Variant, Optional iDirection As Integer = dmVandas.vbColumn) As Variant
    Dim vItems As Variant, vitem As Variant
    Dim sLine As String, sSign As String, sJump As String
    Dim lStart As Long, lEnd As Long, lJump As Long, llimit As Long, lCount As Long
    Dim iDelta As Integer
    
    If Not isInitializedArray(vInterval) Then getInterval = Array(): Exit Function
    
    iDelta = UBound(vInterval) - LBound(vInterval)
    iDelta = iDelta + 1
    If iDelta < 2 Then getInterval = Array(): Exit Function
        
    llimit = IIf(iDirection = dmVandas.vbColumn, column.length, row.length)
    lStart = checkInterval(vInterval(LBound(vInterval)), iDirection)
    lEnd = checkInterval(vInterval(LBound(vInterval) + 1), iDirection)
    If iDelta = 3 Then
        sJump = CStr(vInterval(UBound(vInterval)))
        If IsNumeric(sJump) And InStr(sJump, C_Dot) = 0 And InStr(sJump, C_Comma) = 0 Then
            lJump = CInt(sJump)
        End If
    End If
    If lJump = 0 Then lJump = 1
    
    ' trying to leave without doing the work
    If Not checkIsBetweenBound(lStart, lEnd, llimit) Then _
        getInterval = Array(): Exit Function
    
    ' well, have to do the work
    If lStart < 0 And lJump < 0 Then lJump = lJump * (-1)
    sLine = vbNullString
    sSign = vbNullString
    While lStart <= lEnd
        sLine = sLine & sSign & CStr(lStart)
        lStart = lStart + lJump
        sSign = CSV_SIGN
    Wend
    vItems = IIf(InStr(sLine, CSV_SIGN) > 0, Split(sLine, CSV_SIGN), Array(sLine))
    lCount = LBound(vItems)
    For Each vitem In vItems
        If CLng(vitem) < 0 Then vItems(lCount) = CLng(vitem) * (-1)
        lCount = lCount + 1
    Next vitem

    getInterval = IIf(iDirection = dmVandas.vbColumn, getColumns(vItems), getRows(vItems))
End Function

Private Function checkInterval(vInterval As Variant, Optional ByVal iDirection As Integer = dmVandas.vbline) As Long
    On Error GoTo errorHandling
    Dim vNames As Variant, vName As Variant
    Dim sInterval As String, lPosition As Long
    
    If IsNumeric(vInterval) And InStr(CStr(vInterval), C_Dot) = 0 And _
            InStr(CStr(vInterval), C_Comma) = 0 Then
        checkInterval = CLng(vInterval)
        Exit Function
    End If
    
    vNames = IIf(iDirection = dmVandas.vbline, getLineNames, getColumnNames)
    sInterval = CStr(vInterval)
    lPosition = 1
    For Each vName In vNames
        If CStr(vName) = sInterval Then checkInterval = lPosition: Exit Function
        lPosition = lPosition + 1
    Next vName
    
done:
    Exit Function
errorHandling:
    checkInterval = 0
End Function

Private Function adjustData(ByVal vData As Variant) As Variant
    Dim vResult As Variant, vitem As Variant
    Dim lDimenstions As Long, lSize As Long, lCountRow As Long, lRow As Long, _
        lColumn As Long, lCountColumn As Long
    
    If Not isInitializedArray(vData) Then adjustData = vData: Exit Function
    If Not hasTwoDimentions(vData) Then
        If LBound(vData, 1) = 1 Then adjustData = vData: Exit Function ' no need for adjustment
        lSize = (UBound(vData, 1) - LBound(vData, 1)) + 1
        lCountRow = 1
        ReDim vResult(lCountRow To lSize) As Variant
        For Each vitem In vData
            vResult(lCountRow) = vitem
            lCountRow = lCountRow + 1
        Next vitem
        adjustData = vResult: Exit Function
    End If
    
    lColumn = (UBound(vData, 2) - LBound(vData, 2)) + 1
    lRow = (UBound(vData, 1) - LBound(vData, 1)) + 1
    If LBound(vData, 1) = 1 And LBound(vData, 2) = 1 Then adjustData = vData: Exit Function 'no need for adjustment
    lCountRow = 1
    lCountColumn = 1
    ReDim vResult(lCountRow To lRow, lCountColumn To lColumn) As Variant
    For Each vitem In vData
        vResult(lCountRow, lCountColumn) = vitem
        lCountRow = lCountRow + 1
        If lCountRow > lRow Then _
            lCountColumn = lCountColumn + 1: lCountRow = 1
    Next vitem
    
    adjustData = vResult
End Function

Private Function checkIsBetweenBound(lStart As Long, lEnd As Long, llimit As Long) As Boolean
    If lStart > llimit Or lEnd > llimit Then _
        checkIsBetweenBound = False: Exit Function
    If Abs(lStart) > llimit Or Abs(lEnd) > llimit Then _
        checkIsBetweenBound = False: Exit Function
    If lStart > lEnd Then _
        checkIsBetweenBound = False: Exit Function
    If lStart = 0 Or lEnd = 0 Then _
        checkIsBetweenBound = False: Exit Function
    If lStart < 0 And lEnd > 0 Then _
        checkIsBetweenBound = False: Exit Function
    
    checkIsBetweenBound = True
End Function

Public Function push(ByVal vItems As Variant, Optional ByVal lIndex As Long = 0) As Object
    On Error GoTo errorHandling
    Dim lSizes As Long
    
    
    If Not isInitializedArray(vItems) Then vItems = Array(vItems)
    
    If hasTwoDimentions(vItems) Then
        Call pushMatrix(vItems, lIndex)
    Else
        Call pushLine(vItems, lIndex)
    End If
    
    columnHeaders = vbEmpty
    rowHeaders = vbEmpty
    dimention = dmVandas.vbBoth
    classDimention = dmVandas.vbBoth
    inputDimention = 0
    Set push = Me
done:
    Exit Function
errorHandling:
    columnHeaders = vbEmpty
    rowHeaders = vbEmpty
    dimention = dmVandas.vbBoth
    classDimention = dmVandas.vbBoth
    inputDimention = 0
    Set push = Nothing
End Function

Private Function hasTwoDimentions(ByVal vItems As Variant) As Boolean
    On Error GoTo errorHandling
    Dim lTemp As Long
    
    lTemp = UBound(vItems, 2)
    
    hasTwoDimentions = True
done:
    Exit Function
errorHandling:
    hasTwoDimentions = False
End Function

Private Sub pushLine(ByVal vItems As Variant, Optional ByVal lIndex As Long = 0)
    On Error GoTo errorHandling
    Dim vitem As Variant, vResult As Variant
    Dim lDimentionSize As Long, lStart As Long, lEnd As Long, lDimentions As Long, _
        lRows As Long, lColumns As Long, lNewRows As Long, lNewColumns As Long, _
        lRowCount As Long, lColumnCount As Long, lDimention As Long, lExcess As Long, _
        lRowCountFix As Long, lColumnCountFix As Long, lStartRow As Long, lStartColumn As Long

100    lDimentions = getDimentions
110    lDimention = dimention
120    If lDimentions = 1 And lDimention <> dmVandas.vbColumn Then Call insertLine(vItems, lIndex): Exit Sub
130    classDimention = dmVandas.vbline: lRows = length
140    classDimention = dmVandas.vbColumn: lColumns = length
150    lDimentionSize = IIf(lDimention = dmVandas.vbColumn, lColumns, lRows)
180    lExcess = IIf(lDimention = dmVandas.vbColumn, lIndex - lColumns, lIndex - lRows)
190    If lExcess < 0 Then lExcess = 1
200    lNewRows = IIf(lDimention <> dmVandas.vbColumn, lRows + lExcess, lRows)
210    lNewColumns = IIf(lDimention = dmVandas.vbColumn, lColumns + lExcess, lColumns)
220    If lIndex = 0 Then
230        lIndex = IIf(lDimention = dmVandas.vbColumn, lNewColumns, lNewRows)
240    End If

250    lRowCountFix = 1
260    lColumnCountFix = 1
270    ReDim vResult(lRowCountFix To lNewRows, lColumnCountFix To lNewColumns) As Variant
280    For lRowCount = 1 To lNewRows
290        If lDimention <> dmVandas.vbColumn And lIndex = lRowCount Then lRowCount = lRowCount + 1
300        For lColumnCount = 1 To lNewColumns
310            If lDimention = dmVandas.vbColumn And lIndex = lColumnCount Then lColumnCount = lColumnCount + 1
315            If lColumns > 0 Then
320                vResult(lRowCount, lColumnCount) = vmData(lRowCountFix, lColumnCountFix)
322            Else
324                vResult(lRowCount, lColumnCount) = vmData(lRowCountFix)
326            End If
330            lColumnCountFix = lColumnCountFix + 1
335            If lColumnCountFix > lColumns Then Exit For
340       Next
345        lColumnCountFix = 1
350        lRowCountFix = lRowCountFix + 1
355        If lRowCountFix > lRows Then Exit For
360    Next lRowCount
    
370    Call addNewValues(vItems, lIndex, vResult)
    
440    data = vResult
done:
    Exit Sub
errorHandling:
    Debug.Print "Gone wrong here: " & CStr(Erl)
End Sub

Private Sub insertLine(ByRef vItems As Variant, ByVal lIndex)
    On Error GoTo errorHandling
    Dim vResult As Variant, vitem As Variant, vNewItem As Variant
    Dim lSize As Long, lCount As Long, lExcess As Long, lNewSize As Long, _
        lDelta As Long, lDisplacement As Long
    
    lSize = length
    lDelta = (UBound(vItems) - LBound(vItems)) + 1
    lExcess = lIndex - lSize
    lExcess = IIf(lExcess < 1, lDelta, (lExcess + lDelta) - 1)
    lNewSize = lSize + lExcess
    If lIndex > lSize Then lIndex = (lNewSize - lDelta) + 1
    
    ReDim vResult(1 To lNewSize) As Variant
    lCount = 1
    lDisplacement = 0
    For Each vitem In vmData
        If lCount = lIndex Then lDisplacement = lDelta
        vResult(lCount + lDisplacement) = vmData(lCount)
        lCount = lCount + 1
    Next vitem
    
    lCount = lIndex
    For Each vNewItem In vItems
        vResult(lCount) = vNewItem
        lCount = lCount + 1
    Next vNewItem
    
    data = vResult
    
done:
    Exit Sub
errorHandling:
    Debug.Print "Vandas. insertLine" & vbNewLine & "Line: " & CStr(Erl) & vbNewLine & "Description: " & Err.Description
End Sub

Private Sub addNewValues(ByVal vItems As Variant, ByVal lIndex As Long, ByRef vData As Variant)
    On Error GoTo errorHandling
    Dim vitem As Variant
    Dim lRowCount As Long, lDimention As Long, lColumnCount As Long
    
140    lDimention = dimention
370    lRowCount = IIf(lDimention <> dmVandas.vbColumn, lIndex, 1)
380    lColumnCount = IIf(lDimention = dmVandas.vbColumn, lIndex, 1)
390    For Each vitem In vItems
400        vData(lRowCount, lColumnCount) = vitem
410        lRowCount = IIf(lDimention = dmVandas.vbColumn, lRowCount + 1, lRowCount)
420        lColumnCount = IIf(lDimention <> dmVandas.vbColumn, lColumnCount + 1, lColumnCount)
430    Next vitem

done:
    Exit Sub
errorHandling:
    Debug.Print "Vandas.addNewValues" & vbNewLine & "Line: " & CStr(Erl) & vbNewLine & "Description: " & Err.Description
End Sub

Private Sub pushMatrix(ByVal vItems As Variant, Optional ByVal lIndex As Long = 0)

End Sub

Public Sub printData(Optional ByVal bComplete As Boolean = False)
    On Error GoTo errorHandling
    Dim lColumns As Long, lRows As Long, lRowStart As Long, _
        lColumnStart As Long, lStartColumn As Long, lStartLine As Long, _
        lKeepStartColumn As Long
    Dim sLine As String, sData As String
    Const C_MaxChar As Integer = 20, C_SpaceChar As String = " ", C_BarChar As String = "|", _
        C_HeaderMarker As String = "*"
    
100    If Not isInitialized Then Exit Sub
110    If getDimentions = 1 Then Call printLine: Exit Sub
    
120    classDimention = dmVandas.vbColumn: lColumns = length
130    classDimention = dmVandas.vbline: lRows = length
140    sLine = String(C_MaxChar / 2, C_SpaceChar) & C_BarChar
    
150    For lColumnStart = 1 To lColumns
160        sLine = sLine & Right(String(C_MaxChar, C_SpaceChar) & CStr(lColumnStart), C_MaxChar) & C_BarChar
170    Next lColumnStart
180    Debug.Print sLine
    
190    lStartLine = 1
200    For lRowStart = 1 To lRows
210        sLine = Right(String(C_MaxChar, C_SpaceChar) & _
            CStr(lRowStart), C_MaxChar / 2) & C_BarChar
220        lStartColumn = 1
230        For lColumnStart = 1 To lColumns
240            sData = vbNullString
245            If UCase(TypeName(vmData(lStartLine, lStartColumn))) <> C_ArrayType Then _
                    sData = CStr(vmData(lStartLine, lStartColumn))
250            If lRowStart = 1 Or lColumnStart = 1 Then sData = C_HeaderMarker & sData
260            sLine = sLine & Right(String(C_MaxChar, C_SpaceChar) & _
                sData, C_MaxChar) & C_BarChar
270                lStartColumn = lStartColumn + 1
280        Next lColumnStart
290        Debug.Print sLine
300        lStartLine = lStartLine + 1
310    Next lRowStart
320    Debug.Print "* - Header"
done:
    Exit Sub
errorHandling:
    Debug.Print "Vandas.printData" & vbNewLine & "Line: " & CStr(Erl) & vbNewLine & "Description: " & Err.Description
End Sub

Private Sub printLine(Optional ByVal bComplete As Boolean = False)
    Dim lSize As Long, lCount As Long, lStartPoint As Long
    Dim sLine As String, sData As String
    Const C_MaxChar As Integer = 10, C_SpaceChar As String = " ", C_BarChar As String = "|"
    
    lSize = length
    For lCount = 1 To lSize
        sData = vbNullString
        If UCase(TypeName(vmData(lCount))) <> C_ArrayType Then _
                sData = CStr(vmData(lCount))
        Debug.Print Right(String(C_MaxChar, C_SpaceChar) & CStr(lCount), C_MaxChar) & _
            C_BarChar & C_SpaceChar & sData
        lStartPoint = lStartPoint + 1
    Next lCount
    
End Sub

Public Function pop(Optional ByVal lIndex As Long = 0) As Variant
    Dim vResult As Variant, vTarget As Variant
    
    vTarget = Array(lIndex)
    If lIndex < 1 Then
        vTarget = IIf(dimention = dmVandas.vbColumn, Array(column.length), Array(row.length))
    End If
    
    If dimention = dmVandas.vbColumn Then
        vResult = getColumns(vTarget)
    Else
        vResult = getRows(vTarget)
    End If
    removeFromData (vTarget)
    
    
    columns = vbEmpty
    rows = vbEmpty
    dimention = dmVandas.vbBoth
    classDimention = dmVandas.vbBoth
    inputDimention = 0
    pop = vResult
End Function

'
' The code below is used to get information about the given data
' nothing here is meant to go outside this module
'

Private Sub removeFromData(ByVal vItems As Variant)
    Dim vResult As Variant, vitem As Variant
    Dim lLineCount As Long, lColumnCount As Long, lRows As Long, _
        lColumns As Long, lDelta As Long, lNewRows As Long, lNewColumns As Long, _
        lDimention As Long, lNewLineCount As Long, lNewColumnCount As Long
    
    If Not isInitializedArray(vItems) Then Exit Sub
    
    classDimention = dmVandas.vbline: lRows = length
    classDimention = dmVandas.vbColumn: lColumns = length
    lDelta = UBound(vItems) - LBound(vItems)
    lDelta = lDelta + 1
    lDimention = IIf(dimention = dmVandas.vbColumn, dmVandas.vbColumn, dmVandas.vbline)
    If Not checkItemsToRemove(vItems, lDimention) Then Exit Sub
    lNewRows = IIf(lDimention = dmVandas.vbline, lRows - lDelta, lRows)
    lNewColumns = IIf(lDimention = dmVandas.vbColumn, lColumns - lDelta, lColumns)
    lNewColumnCount = 1
    lNewLineCount = 1

    ReDim vResult(1 To lNewRows, 1 To lNewColumns) As Variant
    For lLineCount = 1 To lRows
        If lDimention = dmVandas.vbline Then
            While getArrayPosition(lLineCount, vItems) > 0
                lLineCount = lLineCount + 1
            Wend
        End If
        If lLineCount > lRows Then Exit For
        For lColumnCount = 1 To lColumns
            If lDimention = dmVandas.vbColumn Then
                While getArrayPosition(lColumnCount, vItems) > 0
                    lColumnCount = lColumnCount + 1
                Wend
            End If
            If lColumnCount > lRows Then Exit For
            vResult(lNewLineCount, lNewColumnCount) = vmData(lLineCount, lColumnCount)
            lNewColumnCount = lNewColumnCount + 1
        Next lColumnCount
        lNewLineCount = lNewLineCount + 1
        lNewColumnCount = 1
    Next lLineCount
    
    data = vResult
    vmColumnNames = vbEmpty
    vmLineNames = vbEmpty
    dimention = dmVandas.vbBoth
End Sub

Private Function checkItemsToRemove(ByRef vItems As Variant, ByRef lDimention As Long) As Boolean
    Dim vitem As Variant
    Dim lCount As Long
    
    lCount = IIf(lDimention = dmVandas.vbline, row.length, column.length)
    For Each vitem In vItems
        If CLng(vitem) < 1 Or CLng(vitem) > lCount Then _
            checkItemsToRemove = False: Exit Function
    Next vitem
    
    checkItemsToRemove = True
End Function

' Maibe add a constant to avoid to loop without exit condition
Private Function getDimentions() As Integer
    On Error GoTo errorHandling
    Dim iTemp As Integer, iResult As Integer
    
        
    If inputDimention > 0 Then getDimentions = inputDimention: Exit Function
    If Not isInitialized Then getDimentions = 0: Exit Function
        
    iResult = 1
    While True
        iTemp = UBound(vmData, iResult)
        iResult = iResult + 1
    Wend

done:
    Exit Function
errorHandling:
    inputDimention = iResult - 1
    getDimentions = iResult - 1
End Function

Private Function lengthMatrix() As Long
    Dim lResultColumns As Long, lResultRows As Long, lResult As Long
    
    lResultColumns = lengthDimention(dmVandas.vbColumn)
    lResultRows = lengthDimention(dmVandas.vbline)
    lResult = IIf(lResultRows > lResultColumns, lResultRows, lResultColumns)
    If lResultColumns > 0 And lResultRows > 0 Then _
        lResult = lResultColumns * lResultRows
    
    lengthMatrix = lResult
End Function

Private Function isInitializedArray(ByRef vArray As Variant) As Boolean
    On Error GoTo errorHandling
    Dim bResult As Boolean
    Dim lDimention As Long
    
    bResult = False
    If IsNumeric(UBound(vArray)) Then bResult = True
    If UBound(vArray) < 0 Then bResult = False
    
    isInitializedArray = bResult
done:
    Exit Function
errorHandling:
    isInitializedArray = bResult
End Function

Private Function isInitialized() As Boolean
    isInitialized = isInitializedArray(data)
End Function


' End block


