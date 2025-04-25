'==============================================================================
' Module: SheetUtilities
' Author: Kyle Hsu
' Date: April 21, 2025
' Description: Excel sheet management utility functions
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Function: SheetExists
' Purpose: Checks if a worksheet with the specified name exists in the workbook
' Parameters: strSheetName - Name of the sheet to check
' Returns: Boolean - True if the sheet exists, False otherwise
'------------------------------------------------------------------------------
Public Function SheetExists(strSheetName As String) As Boolean
    Dim ws As Worksheet
    
    SheetExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strSheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

'------------------------------------------------------------------------------
' Sub: CreateOrReplaceSheet
' Purpose: Creates a new worksheet or replaces an existing one with the same name
' Parameters:
'   strSheetName - Name for the sheet
'   objPosition - Worksheet to position the new sheet after (optional)
'   blnPromptBeforeReplace - Whether to prompt before replacing (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub CreateOrReplaceSheet(strSheetName As String, _
                               Optional objPosition As Worksheet = Nothing, _
                               Optional blnPromptBeforeReplace As Boolean = False)
    Dim blnReplaceSheet As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Check if sheet already exists
    If SheetExists(strSheetName) Then
        If blnPromptBeforeReplace Then
            blnReplaceSheet = (MsgBox("Sheet '" & strSheetName & "' already exists. Replace it?", _
                              vbQuestion + vbYesNo, "Confirm Sheet Replacement") = vbYes)
            
            If Not blnReplaceSheet Then
                Exit Sub
            End If
        End If
        
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(strSheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new worksheet
    If objPosition Is Nothing Then
        ' Add at the end
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = strSheetName
    Else
        ' Add after the specified worksheet
        ThisWorkbook.Worksheets.Add(After:=objPosition).Name = strSheetName
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in CreateOrReplaceSheet", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: HideUnusedRows
' Purpose: Hides unused rows based on a specified condition column
' Parameters:
'   strSheetName - Name of the sheet to process
'   strConditionColumn - Column to check for empty/zero values
'   intStartRow - First row to check
'   intEndRow - Last row to check
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub HideUnusedRows(strSheetName As String, _
                         strConditionColumn As String, _
                         intStartRow As Integer, _
                         intEndRow As Integer)
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Worksheets(strSheetName).Activate
    
    For i = intStartRow To intEndRow
        If WorksheetFunction.CountA(Range(strConditionColumn & i)) = 0 Or _
           Val(Range(strConditionColumn & i).Value) = 0 Then
            Rows(i & ":" & i).EntireRow.Hidden = True
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in HideUnusedRows", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: CopyRangeAcrossSheets
' Purpose: Copies a range from one sheet to multiple destination sheets
' Parameters:
'   objSourceRange - Range to copy
'   arrDestSheets - Array of sheet names to copy to
'   strDestRange - Destination cell reference (top-left corner)
'   blnPasteValues - Whether to paste values only (True) or formulas (False)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub CopyRangeAcrossSheets(objSourceRange As Range, _
                                arrDestSheets() As String, _
                                strDestRange As String, _
                                Optional blnPasteValues As Boolean = True)
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    objSourceRange.Copy
    
    For i = LBound(arrDestSheets) To UBound(arrDestSheets)
        If SheetExists(arrDestSheets(i)) Then
            ThisWorkbook.Worksheets(arrDestSheets(i)).Range(strDestRange).Select
            
            If blnPasteValues Then
                Selection.PasteSpecial Paste:=xlPasteValues
            Else
                Selection.PasteSpecial Paste:=xlPasteFormulas
            End If
        End If
    Next i
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in CopyRangeAcrossSheets", vbCritical, "Error"
End Sub
