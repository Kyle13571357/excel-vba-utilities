'==============================================================================
' Module: UIEnhancements
' Author: Kyle Hsu
' Date: April 21, 2025
' Description: Excel UI enhancement utilities for buttons, formatting, and styles
'==============================================================================
Option Explicit

' Color constants for consistent styling
Public Const COLOR_PRIMARY As Long = &H36009A   ' RGB(154, 0, 54)
Public Const COLOR_SECONDARY As Long = &H50B000 ' RGB(0, 176, 80)
Public Const COLOR_ACCENT As Long = &HF69D00    ' RGB(0, 157, 246)
Public Const COLOR_WARNING As Long = &H0050FF   ' RGB(255, 80, 0)
Public Const COLOR_SUCCESS As Long = &H7CFC00   ' RGB(0, 252, 124)

'------------------------------------------------------------------------------
' Sub: CreateCommandButton
' Purpose: Creates a formatted command button shape with standardized styling
' Parameters:
'   strSheetName - Sheet where button should be placed
'   strButtonCaption - Text to display on button
'   strMacroName - Name of macro to run when clicked
'   dblLeft - Left position (in points)
'   dblTop - Top position (in points)
'   dblWidth - Width of button (in points)
'   dblHeight - Height of button (in points)
'   lngButtonColor - Background color (optional)
'   lngTextColor - Text color (optional)
' Returns: Shape - Reference to the created button shape
'------------------------------------------------------------------------------
Public Function CreateCommandButton(strSheetName As String, _
                                  strButtonCaption As String, _
                                  strMacroName As String, _
                                  dblLeft As Double, _
                                  dblTop As Double, _
                                  dblWidth As Double, _
                                  dblHeight As Double, _
                                  Optional lngButtonColor As Long = &H36009A, _
                                  Optional lngTextColor As Long = &HFFFFFF) As Shape
    Dim objSheet As Worksheet
    Dim objButton As Shape
    
    On Error GoTo ErrorHandler
    
    ' Activate the target sheet
    Set objSheet = ThisWorkbook.Worksheets(strSheetName)
    objSheet.Activate
    
    ' Create rounded rectangle shape
    Set objButton = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                                                dblLeft, dblTop, dblWidth, dblHeight)
    
    ' Format button appearance
    With objButton
        .Line.Visible = msoFalse
        
        ' Set 3D effect
        With .ThreeD
            .SetPresetCamera (msoCameraOrthographicFront)
            .LightAngle = 145
            .BevelTopInset = 12
            .BevelTopDepth = 3
            .BevelBottomType = msoBevelNone
        End With
        
        ' Set text
        .TextFrame2.TextRange.Characters.Text = strButtonCaption
        
        ' Format text
        With .TextFrame2.TextRange.Font
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = lngTextColor
            .Bold = msoTrue
            .Size = 14
            .Name = "Calibri"
        End With
        
        ' Set fill color
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = lngButtonColor
        .Fill.Transparency = 0
        
        ' Set action
        .OnAction = strMacroName
    End With
    
    Set CreateCommandButton = objButton
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in CreateCommandButton", vbCritical, "Error"
    Set CreateCommandButton = Nothing
End Function

'------------------------------------------------------------------------------
' Sub: CreateButtonPanel
' Purpose: Creates a panel with multiple command buttons
' Parameters:
'   strSheetName - Sheet where button panel should be placed
'   arrButtonInfo - Array of button information (caption, macro, color)
'   dblStartLeft - Left position for first button
'   dblStartTop - Top position for first button
'   dblButtonWidth - Width of each button
'   dblButtonHeight - Height of each button
'   dblHorizontalSpacing - Horizontal space between buttons
'   dblVerticalSpacing - Vertical space between rows of buttons
'   intButtonsPerRow - Number of buttons per row
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub CreateButtonPanel(strSheetName As String, _
                           arrButtonInfo As Variant, _
                           dblStartLeft As Double, _
                           dblStartTop As Double, _
                           dblButtonWidth As Double, _
                           dblButtonHeight As Double, _
                           dblHorizontalSpacing As Double, _
                           dblVerticalSpacing As Double, _
                           intButtonsPerRow As Integer)
    Dim i As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrorHandler
    
    ' Activate target sheet
    ThisWorkbook.Worksheets(strSheetName).Activate
    
    ' Loop through button information and create buttons
    For i = LBound(arrButtonInfo) To UBound(arrButtonInfo)
        ' Calculate button position
        intRow = (i - LBound(arrButtonInfo)) \ intButtonsPerRow
        intCol = (i - LBound(arrButtonInfo)) Mod intButtonsPerRow
        
        dblLeft = dblStartLeft + intCol * (dblButtonWidth + dblHorizontalSpacing)
        dblTop = dblStartTop + intRow * (dblButtonHeight + dblVerticalSpacing)
        
        ' Create button
        CreateCommandButton strSheetName, _
                         arrButtonInfo(i)(0), _  ' Button caption
                         arrButtonInfo(i)(1), _  ' Macro name
                         dblLeft, _
                         dblTop, _
                         dblButtonWidth, _
                         dblButtonHeight, _
                         arrButtonInfo(i)(2)     ' Button color
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in CreateButtonPanel", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyStandardCellFormatting
' Purpose: Applies standardized formatting to a range of cells
' Parameters:
'   objRange - Range to format
'   strFontName - Font name
'   intFontSize - Font size
'   blnBold - Whether text is bold
'   lngFontColor - Font color
'   lngFillColor - Cell background color
'   blnShrinkToFit - Whether to shrink text to fit cell
'   intHorizontalAlignment - Horizontal alignment
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyStandardCellFormatting(objRange As Range, _
                                     Optional strFontName As String = "Calibri", _
                                     Optional intFontSize As Integer = 11, _
                                     Optional blnBold As Boolean = False, _
                                     Optional lngFontColor As Long = vbBlack, _
                                     Optional lngFillColor As Long = xlNone, _
                                     Optional blnShrinkToFit As Boolean = False, _
                                     Optional intHorizontalAlignment As Integer = xlLeft)
    On Error GoTo ErrorHandler
    
    With objRange
        ' Font properties
        .Font.Name = strFontName
        .Font.Size = intFontSize
        .Font.Bold = blnBold
        .Font.Color = lngFontColor
        
        ' Cell properties
        If lngFillColor <> xlNone Then
            .Interior.Color = lngFillColor
        End If
        
        .HorizontalAlignment = intHorizontalAlignment
        .VerticalAlignment = xlCenter
        .WrapText = False
        .ShrinkToFit = blnShrinkToFit
        .ReadingOrder = xlContext
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyStandardCellFormatting", vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyStandardBorders
' Purpose: Applies standardized borders to a range of cells
' Parameters:
'   objRange - Range to apply borders to
'   blnOutsideBorders - Whether to apply outside borders
'   blnInsideBorders - Whether to apply inside borders
'   lngOutsideColor - Color for outside borders
'   lngInsideColor - Color for inside borders
'   intOutsideWeight - Weight for outside borders
'   intInsideWeight - Weight for inside borders
' Returns: Nothing
'------------------------------------------------------------------------------
Public Sub ApplyStandardBorders(objRange As Range, _
                              Optional blnOutsideBorders As Boolean = True, _
                              Optional blnInsideBorders As Boolean = False, _
                              Optional lngOutsideColor As Long = vbBlack, _
                              Optional lngInsideColor As Long = vbBlack, _
                              Optional intOutsideWeight As Integer = xlMedium, _
                              Optional intInsideWeight As Integer = xlThin)
    On Error GoTo ErrorHandler
    
    ' Clear existing borders
    objRange.Borders(xlDiagonalDown).LineStyle = xlNone
    objRange.Borders(xlDiagonalUp).LineStyle = xlNone
    
    ' Apply outside borders if requested
    If blnOutsideBorders Then
        With objRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = lngOutsideColor
            .Weight = intOutsideWeight
        End With
        
        With objRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = lngOutsideColor
            .Weight = intOutsideWeight
        End With
        
        With objRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = lngOutsideColor
            .Weight = intOutsideWeight
        End With
        
        With objRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = lngOutsideColor
            .Weight = intOutsideWeight
        End With
    End If
    
    ' Apply inside borders if requested
    If blnInsideBorders Then
        With objRange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Color = lngInsideColor
            .Weight = intInsideWeight
        End With
        
        With objRange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Color = lngInsideColor
            .Weight = intInsideWeight
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ApplyStandardBorders", vbCritical, "Error"
End Sub
