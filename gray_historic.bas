Attribute VB_Name = "Module1"
' ==============================================================================
' FillHistoricalGray
'
' Purpose : Marks historical columns in a financial model by applying a gray
'           fill to every cell in the selection that currently has either no
'           fill (transparent/blank) or a white fill.
'           Cells that already carry any other color (blue headers, black
'           labels, etc.) are intentionally left untouched.
'
' How to use:
'   1. Select the range of cells you want to process.
'   2. Run this macro (e.g. via Alt+F8 > FillHistoricalGray > Run).
'
' Paste this code into PERSONAL.XLSB > Module1 so it is available in every
' workbook.
' ==============================================================================

Sub FillHistoricalGray()

    Dim rng  As Range
    Dim cell As Range
    Dim ci   As Long

    ' ---- Color constants -------------------------------------------------------
    ' Historical gray: RGB(217, 217, 217) - Excel built-in "Gray 25%"
    Const HISTORICAL_GRAY As Long = 14277081
    ' White: RGB(255, 255, 255)
    Const WHITE_COLOR     As Long = 16777215
    ' ---------------------------------------------------------------------------

    ' Guard: a range must be selected before calling the macro
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells before running this macro.", _
               vbExclamation, "FillHistoricalGray - No Range Selected"
        Exit Sub
    End If

    Set rng = Selection

    Application.ScreenUpdating = False

    For Each cell In rng

        ci = cell.Interior.ColorIndex

        ' xlColorIndexNone (-4142) means the cell has no background fill at all.
        ' We also catch explicit white fills by comparing the Color property.
        ' Every other color (blue, black, yellow, …) is skipped automatically.
        If ci = xlColorIndexNone Or cell.Interior.Color = WHITE_COLOR Then
            cell.Interior.Color = HISTORICAL_GRAY
        End If

    Next cell

    Application.ScreenUpdating = True

End Sub
