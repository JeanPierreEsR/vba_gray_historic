Attribute VB_Name = "Module2"
' ==============================================================================
' ClearHistoricalGray
'
' Purpose : Reverses the effect of FillHistoricalGray by removing the gray
'           fill from every cell in the selection that carries exactly the
'           historical-gray color – RGB(242, 242, 242) / Excel "Gray 10%".
'           All other fills (blue headers, black labels, white, etc.) are
'           intentionally left untouched.
'
' How to use:
'   1. Select the range of cells you want to process.
'   2. Run this macro (e.g. via Alt+F8 > ClearHistoricalGray > Run).
'
' Paste this code into PERSONAL.XLSB > Module2 so it is available in every
' workbook.
' ==============================================================================

Sub ClearHistoricalGray()

    Dim rng  As Range
    Dim cell As Range

    ' ---- Color constant --------------------------------------------------------
    ' Historical gray: RGB(242, 242, 242) - Excel built-in "Gray 10%"
    Const HISTORICAL_GRAY As Long = 15921906
    ' ---------------------------------------------------------------------------

    ' Guard: a range must be selected before calling the macro
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells before running this macro.", _
               vbExclamation, "ClearHistoricalGray - No Range Selected"
        Exit Sub
    End If

    Set rng = Selection

    Application.ScreenUpdating = False

    For Each cell In rng

        ' Only touch cells that carry exactly the historical-gray fill.
        ' Every other color (white, blue, black, yellow, …) is skipped.
        If cell.Interior.ColorIndex <> xlColorIndexNone And _
           cell.Interior.Color = HISTORICAL_GRAY Then
            cell.Interior.ColorIndex = xlColorIndexNone
        End If

    Next cell

    Application.ScreenUpdating = True

End Sub
