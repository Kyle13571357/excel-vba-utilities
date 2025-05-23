'==============================================================================
' Module: ExportUtilities
' Author: Kyle Hsu
' Date: April 21, 2025
' Description: Excel workbook export utilities for PNG and PDF formats
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: ExportSheetAsPNG
' Purpose: Exports a worksheet range as a PNG image
' Parameters:
'   objRangeToExport - Range to export as image
'   strExportPath - Path where the image should be saved
'   strFileName - Base filename (without extension)
'   blnAddTimestamp - Whether to add timestamp to filename
'   dblScaleFactor - Scale factor for the exported image
' Returns: String - Full path of the exported file
'------------------------------------------------------------------------------
Public Function ExportSheetAsPNG(objRangeToExport As Range, _
                               strExportPath As String, _
                               strFileName As String, _
                               Optional blnAddTimestamp As Boolean = True, _
                               Optional dblScaleFactor As Double = 2) As String
    Dim strFullPath As String
    Dim strTimestamp As String
    Dim objChart As ChartObject
    
    On Error GoTo ErrorHandler
    
    ' Create export directory if it doesn't exist
    If Len(Dir(strExportPath, vbDirectory)) = 0 Then
        MkDir strExportPath
    End If
    
    ' Generate timestamp if needed
    If blnAddTimestamp Then
        strTimestamp = Format(Now, "yyyymmdd_hhmmss")
    Else
        strTimestamp = ""
    End If
    
    ' Build full path
    If Right(strExportPath, 1) <> "\" Then
        strExportPath = strExportPath & "\"
    End If
    
    strFullPath = strExportPath & strFileName & IIf(strTimestamp = "", "", "_" & strTimestamp) & ".png"
    
    ' Create temporary chart object for exporting
    Application.ScreenUpdating = False
    
    ' Copy the range as a picture
    objRangeToExport.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    ' Create a temporary chart to paste the picture
    Set objChart = ActiveSheet.ChartObjects.Add( _
        Left:=0, _
        Top:=0, _
        Width:=objRangeToExport.Width * dblScaleFactor, _
        Height:=objRangeToExport.Height * dblScaleFactor)
    
    ' Paste and export as PNG
    objChart.Chart.Paste
    objChart.Chart.Export strFullPath, "PNG"
    
    ' Clean up
    objChart.Delete
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    ' Return the full path
    ExportSheetAsPNG = strFullPath
    
    Exit Function
    
ErrorHandler:
    If Not objChart Is Nothing Then
        objChart.Delete
    End If
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ExportSheetAsPNG", vbCritical, "Error"
    ExportSheetAsPNG = ""
End Function

'------------------------------------------------------------------------------
' Sub: ExportSheetsAsPDF
' Purpose: Exports multiple worksheets as a single PDF file
' Parameters:
'   arrSheets - Array of sheet names to export
'   strExportPath - Path where the PDF should be saved
'   strFileName - Base filename (without extension)
'   blnAddTimestamp - Whether to add timestamp to filename
'   blnOpenAfterExport - Whether to open the PDF after export
' Returns: String - Full path of the exported file
'------------------------------------------------------------------------------
Public Function ExportSheetsAsPDF(arrSheets() As String, _
                                strExportPath As String, _
                                strFileName As String, _
                                Optional blnAddTimestamp As Boolean = True, _
                                Optional blnOpenAfterExport As Boolean = False) As String
    Dim strFullPath As String
    Dim strTimestamp As String
    Dim objSheets As Object
    Dim i As Integer
    Dim blnSheetsExist As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Verify sheets exist
    blnSheetsExist = True
    For i = LBound(arrSheets) To UBound(arrSheets)
        If Not SheetExists(arrSheets(i)) Then
            blnSheetsExist = False
            MsgBox "Sheet '" & arrSheets(i) & "' does not exist.", vbExclamation, "Export Error"
            Exit Function
        End If
    Next i
    
    ' Create export directory if it doesn't exist
    If Len(Dir(strExportPath, vbDirectory)) = 0 Then
        MkDir strExportPath
    End If
    
    ' Generate timestamp if needed
    If blnAddTimestamp Then
        strTimestamp = Format(Now, "yyyymmdd_hhmmss")
    Else
        strTimestamp = ""
    End If
    
    ' Build full path
    If Right(strExportPath, 1) <> "\" Then
        strExportPath = strExportPath & "\"
    End If
    
    strFullPath = strExportPath & strFileName & IIf(strTimestamp = "", "", "_" & strTimestamp) & ".pdf"
    
    ' Select sheets to export
    Sheets(arrSheets).Select
    
    ' Export as PDF
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=strFullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=blnOpenAfterExport
    
    ' Return to first sheet and deselect multiple sheets
    ThisWorkbook.Worksheets(1).Select
    
    ' Return the full path
    ExportSheetsAsPDF = strFullPath
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ExportSheetsAsPDF", vbCritical, "Error"
    ExportSheetsAsPDF = ""
End Function

'------------------------------------------------------------------------------
' Function: SheetExists
' Purpose: Helper function to check if a worksheet exists
' Parameters: strSheetName - Name of sheet to check
' Returns: Boolean - Whether the sheet exists
'------------------------------------------------------------------------------
Private Function SheetExists(strSheetName As String) As Boolean
    Dim ws As Worksheet
    
    SheetExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strSheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
