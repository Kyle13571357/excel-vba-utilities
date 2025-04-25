'==============================================================================
' Module: IntervalValidationUtilities
' Purpose: Validate and rebuild shift schedule compliance logic
' Author: Kyle Hsu
' Date: 2025/04/21
'==============================================================================

Option Explicit

'==================== 常數區（集中配置工作表與範圍） ===================='
Private Const SHEET_T1_CHECK As String = "T1間隔時間CHECK"
Private Const SHEET_T2_CHECK As String = "T2間隔時間CHECK"
Private Const SHEET_T1_SOURCE As String = "T1輪 "
Private Const SHEET_T2_SOURCE As String = "T2輪"
Private Const SHEET_REF As String = "T1助櫃"
Private Const T1_LOOKUP_RANGE As String = "$A$192:$BW$196"
Private Const T2_LOOKUP_RANGE As String = "$A$203:$BW$206"
Private Const REF_CELL_T1 As String = "B10"
Private Const MAX_ROW_T1 As Long = 183
Private Const MAX_ROW_T2 As Long = 199
Private Const FORMULA_STEP As Long = 4

'==============================================================================
' Entry Point: ResetIntervalCheckFormulas
'==============================================================================
Public Sub ResetIntervalCheckFormulas()
    Dim blnProceed As Boolean

    Sheets(SHEET_T1_CHECK).Select
    If Range("D3") = "六" Then
        MsgBox "請確認是否已執行過此功能，勿重複執行", vbOKOnly, "警告"
        blnProceed = (MsgBox("請確認是否執行", vbYesNo, "Warning") = vbYes)
    Else
        blnProceed = True
    End If

    If blnProceed Then
        Call FormatCheckSheet(SHEET_T1_CHECK, SHEET_T1_SOURCE, SHEET_REF, T1_LOOKUP_RANGE, MAX_ROW_T1)
        Call FormatCheckSheet(SHEET_T2_CHECK, SHEET_T2_SOURCE, SHEET_REF, T2_LOOKUP_RANGE, MAX_ROW_T2)
    End If
End Sub

'==============================================================================
Private Sub FormatCheckSheet(strSheetName As String, strSourceSheet As String, _
                           strRefSheet As String, strLookupRange As String, maxRow As Long)
    Dim i As Long

    Sheets(strSheetName).Select
    If strSheetName = SHEET_T1_CHECK Then
        Range("B180") = "=" & strRefSheet & "!" & REF_CELL_T1
    End If

    Call InsertFormatColumns(strSheetName)
    Range("D2") = Sheets(strSourceSheet).Range("AP61")
    Range("D3") = Sheets(strSourceSheet).Range("AP62")

    Range("D4") = "=IF(ISNA(VLOOKUP(B4,'" & SHEET_T1_SOURCE & "'!$AK$63:$AP$107,6,FALSE))," & _
                     "VLOOKUP(B4," & SHEET_T2_SOURCE & "!$AK$63:$AP$111,6,FALSE)," & _
                     "VLOOKUP(B4,'" & SHEET_T1_SOURCE & "'!$AK$63:$AP$107,6,FALSE))"

    Range("D5") = "=IF(ISNA(HLOOKUP(D4," & strLookupRange & ",3,FALSE)),\"\",(HLOOKUP(D4," & strLookupRange & ",3,FALSE)))"
    Range("D6") = "=IF(ISNA(HLOOKUP(D4," & strLookupRange & ",4,FALSE)),\"\",(HLOOKUP(D4," & strLookupRange & ",4,FALSE)))"
    Range("D7") = "=IF(F5=\"\",\"\",IF(D6=\"\",\"\",IF(D6<$F$190,F5-D6,(F5-D6+1))))"

    Range("D4:E7").Select
    Selection.Copy
    i = 8
    Do Until i > maxRow
        Range("D" & i).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas
        i = i + FORMULA_STEP
    Loop
    Application.CutCopyMode = False
End Sub

'==============================================================================
Private Sub InsertFormatColumns(strSheetName As String)
    Sheets(strSheetName).Select

    If strSheetName = SHEET_T1_CHECK Then
        Range("D2:E183").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ClearBorderStyles Range("D2:E183")
        Range("F2:G183").Copy
        Range("D2").PasteSpecial Paste:=xlPasteFormats
        SetLeftBorder Range("D2:E183")
    Else
        Rows("185:188").Insert Shift:=xlDown
        Rows("185:188").Insert Shift:=xlDown
        Rows("196:198").Insert Shift:=xlDown

        Range("D2:E199").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ClearBorderStyles Range("D2:E199")
        Range("F2:G199").Copy
        Range("D2").PasteSpecial Paste:=xlPasteFormats
        SetLeftBorder Range("D2:E199")
    End If
    Application.CutCopyMode = False
End Sub

'==============================================================================
Private Sub ClearBorderStyles(objRange As Range)
    objRange.Borders(xlDiagonalDown).LineStyle = xlNone
    objRange.Borders(xlDiagonalUp).LineStyle = xlNone
    objRange.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

'==============================================================================
Private Sub SetLeftBorder(objRange As Range)
    With objRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub
