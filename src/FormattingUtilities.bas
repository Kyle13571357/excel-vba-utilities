'==============================================================================
' Module: FormattingUtilities
' Author: Your Name Here
' Date: April 21, 2025
' Description: Utilities for consistent formatting across Excel worksheets
'==============================================================================
Option Explicit

' Color constants for standardized formatting
Public Const COLOR_PRIMARY As Long = &H36009A      ' RGB(154, 0, 54)
Public Const COLOR_SECONDARY As Long = &HD9D9D9    ' RGB(217, 217, 217)
Public Const COLOR_TEXT_DARK As Long = &H000000    ' RGB(0, 0, 0)
Public Const COLOR_TEXT_LIGHT As Long = &HFFFFFF   ' RGB(255, 255, 255)
Public Const COLOR_HIGHLIGHT_1 As Long = &HC5C9E7  ' RGB(231, 201, 197)
Public Const COLOR_HIGHLIGHT_2 As Long = &HE1D8BC  ' RGB(188, 216, 225)
Public Const COLOR_WEEKEND As Long = &HFFFF00      ' RGB(0, 255, 255)

'------------------------------------------------------------------------------
' Sub: ApplyHeaderFormatting
' Purpose: Applies standardized header formatting to a range
' Parameters:
'   objRange - Range to format
'   strTitle - Title text (optional)
'   lngBackColor - Background color (optional)
'   lngTextColor - Text color (optional)
'   intFontSize - Font size (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyHeaderFormatting(objRange As Range, _
                               Optional strTitle As String = "", _
                               Optional lngBackColor As Long = -1, _
                               Optional lngTextColor As Long = -1, _
                               Optional intFontSize As Integer = 16)
    On Error GoTo ErrorHandler
    
    ' Use default colors if not specified
    If lngBackColor = -1 Then lngBackColor = COLOR_PRIMARY
    If lngTextColor = -1 Then lngTextColor = COLOR_TEXT_LIGHT
    
    ' Format range
    With objRange
        ' Merge cells if range is multi-cell
        If .Cells.Count > 1 Then
            .MergeCells = True
        End If
        
        ' Set text if provided
        If Len(strTitle) > 0 Then
            .Value = strTitle
        End If
        
        ' Apply formatting
        .Interior.Color = lngBackColor
        
        With .Font
            .Name = "微軟正黑體"
            .Size = intFontSize
            .Bold = True
            .Color = lngTextColor
        End With
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ShrinkToFit = True
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyHeaderFormatting", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyAlternatingRowColors
' Purpose: Applies alternating row colors to a range
' Parameters:
'   objRange - Range to format
'   lngColor1 - First color (optional)
'   lngColor2 - Second color (optional)
'   intStartRow - Row to start alternating pattern (optional)
'   intRowsPerPattern - Number of rows for each color (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyAlternatingRowColors(objRange As Range, _
                                   Optional lngColor1 As Long = -1, _
                                   Optional lngColor2 As Long = -1, _
                                   Optional intStartRow As Integer = 1, _
                                   Optional intRowsPerPattern As Integer = 1)
    Dim i As Long
    Dim j As Long
    Dim intLastRow As Long
    Dim intCurrentRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Use default colors if not specified
    If lngColor1 = -1 Then lngColor1 = COLOR_HIGHLIGHT_1
    If lngColor2 = -1 Then lngColor2 = xlNone
    
    ' Get last row in range
    intLastRow = objRange.Rows.Count
    
    ' Apply alternating colors
    For i = intStartRow To intLastRow Step intRowsPerPattern * 2
        ' First color
        intCurrentRow = i
        For j = 1 To intRowsPerPattern
            If intCurrentRow <= intLastRow Then
                objRange.Rows(intCurrentRow).Interior.Color = lngColor1
                intCurrentRow = intCurrentRow + 1
            End If
        Next j
        
        ' Second color
        For j = 1 To intRowsPerPattern
            If intCurrentRow <= intLastRow Then
                If lngColor2 <> xlNone Then
                    objRange.Rows(intCurrentRow).Interior.Color = lngColor2
                End If
                intCurrentRow = intCurrentRow + 1
            End If
        Next j
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyAlternatingRowColors", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: StandardizeColumnWidths
' Purpose: Applies standardized column widths to a range
' Parameters:
'   objRange - Range to format
'   dblWidth - Width to apply (optional)
'   arrCustomWidths - Array of custom widths for specific columns (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub StandardizeColumnWidths(objRange As Range, _
                                 Optional dblWidth As Double = 0, _
                                 Optional arrCustomWidths As Variant = Null)
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    ' Set standard width if specified
    If dblWidth > 0 Then
        objRange.ColumnWidth = dblWidth
    End If
    
    ' Apply custom widths if specified
    If Not IsNull(arrCustomWidths) Then
        For i = LBound(arrCustomWidths) To UBound(arrCustomWidths)
            j = i - LBound(arrCustomWidths) + 1
            
            If j <= objRange.Columns.Count Then
                objRange.Columns(j).ColumnWidth = arrCustomWidths(i)
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in StandardizeColumnWidths", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyDateColumnFormatting
' Purpose: Highlights weekend days in date columns
' Parameters:
'   objDateRow - Row containing dates
'   lngWeekendColor - Color for weekend cells (optional)
'   objRangeToFormat - Range to apply formatting to (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyDateColumnFormatting(objDateRow As Range, _
                                   Optional lngWeekendColor As Long = -1, _
                                   Optional objRangeToFormat As Range = Nothing)
    Dim i As Long
    Dim datDate As Date
    Dim intDayOfWeek As Integer
    Dim objColumn As Range
    
    On Error GoTo ErrorHandler
    
    ' Use default color if not specified
    If lngWeekendColor = -1 Then lngWeekendColor = COLOR_WEEKEND
    
    ' Process each cell in date row
    For i = 1 To objDateRow.Columns.Count
        On Error Resume Next
        datDate = objDateRow.Cells(1, i).Value
        On Error GoTo ErrorHandler
        
        ' Skip if not a valid date
        If IsDate(datDate) Then
            intDayOfWeek = Weekday(datDate)
            
            ' If weekend (1 = Sunday, 7 = Saturday)
            If intDayOfWeek = 1 Or intDayOfWeek = 7 Then
                ' Format just the date cell if no range specified
                If objRangeToFormat Is Nothing Then
                    objDateRow.Cells(1, i).Interior.Color = lngWeekendColor
                ' Format the entire column in the specified range
                Else
                    Set objColumn = objRangeToFormat.Columns(i)
                    objColumn.Interior.Color = lngWeekendColor
                End If
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyDateColumnFormatting", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyConditionalFormatting
' Purpose: Applies conditional formatting to highlight specific values
' Parameters:
'   objRange - Range to format
'   arrValuesToHighlight - Array of values to highlight
'   lngHighlightColor - Color for highlighted cells
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyConditionalFormatting(objRange As Range, _
                                    arrValuesToHighlight As Variant, _
                                    lngHighlightColor As Long)
    Dim i As Long
    Dim strValue As String
    
    On Error GoTo ErrorHandler
    
    ' Clear existing conditional formatting
    objRange.FormatConditions.Delete
    
    ' Add condition for each value
    For i = LBound(arrValuesToHighlight) To UBound(arrValuesToHighlight)
        strValue = CStr(arrValuesToHighlight(i))
        
        objRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & strValue & """"
        objRange.FormatConditions(i + 1).Interior.Color = lngHighlightColor
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyConditionalFormatting", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: FormatAsTable
' Purpose: Applies consistent table formatting to a range
' Parameters:
'   objRange - Range to format
'   lngHeaderColor - Color for header row (optional)
'   lngBodyColor1 - First alternating color (optional)
'   lngBodyColor2 - Second alternating color (optional)
'   blnBoldHeaderRow - Whether to bold the header row (optional)
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub FormatAsTable(objRange As Range, _
                       Optional lngHeaderColor As Long = -1, _
                       Optional lngBodyColor1 As Long = -1, _
                       Optional lngBodyColor2 As Long = -1, _
                       Optional blnBoldHeaderRow As Boolean = True)
    On Error GoTo ErrorHandler
    
    ' Use default colors if not specified
    If lngHeaderColor = -1 Then lngHeaderColor = COLOR_PRIMARY
    If lngBodyColor1 = -1 Then lngBodyColor1 = COLOR_HIGHLIGHT_1
    If lngBodyColor2 = -1 Then lngBodyColor2 = RGB(255, 255, 255)
    
    ' Format header row
    With objRange.Rows(1)
        .Interior.Color = lngHeaderColor
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = blnBoldHeaderRow
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Format body rows with alternating colors
    If objRange.Rows.Count > 1 Then
        ApplyAlternatingRowColors _
            Range(objRange.Rows(2), objRange.Rows(objRange.Rows.Count)), _
            lngBodyColor1, _
            lngBodyColor2, _
            1, _
            1
    End If
    
    ' Add borders
    With objRange.Borders
        .LineStyle = xlContinuous
        .Color = COLOR_TEXT_DARK
        .Weight = xlThin
    End With
    
    With objRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = COLOR_TEXT_DARK
        .Weight = xlMedium
    End With
    
    With objRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = COLOR_TEXT_DARK
        .Weight = xlMedium
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in FormatAsTable", vbCritical, "Error"
End Sub
