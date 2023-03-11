'*******************************************************************************
'*
'*  Author: Sven from CodingIsFun
'*  Website: https://pythonandvba.com
'*  YouTube: https://youtube.com/@codingisfun
'*
'*  This module contains two macros:
'*     - IfErrorBlank: Adds an IFERROR function to selected cells and returns a blank cell if an error occurs
'*     - IfErrorZero: Adds an IFERROR function to selected cells and returns zero if an error occurs
'*
'*******************************************************************************

Option Explicit

Public Sub IfErrorBlank(control As IRibbonControl)
    'Add an IFERROR function to all selected cells with a formula and return a blank cell if an error occurs

    Dim cell As Range
    
    On Error Resume Next
    For Each cell In Selection.Cells
        If cell.HasFormula And Not cell.HasArray Then
            'Add the IFERROR function to the formula
            cell.formula = AddIfError(cell.formula, """""")
        End If
    Next cell
    On Error GoTo 0
End Sub

Public Sub IfErrorZero(control As IRibbonControl)
    'Add an IFERROR function to all selected cells with a formula and return zero if an error occurs

    Dim cell As Range
    
    On Error Resume Next
    For Each cell In Selection.Cells
        If cell.HasFormula And Not cell.HasArray Then
            'Add the IFERROR function to the formula
            cell.formula = AddIfError(cell.formula, 0)
        End If
    Next cell
    On Error GoTo 0
End Sub

Private Function AddIfError(ByVal formula As String, ByVal errorValue As Variant) As String
    'Adds an IFERROR function to the formula and returns the specified error value if an error occurs
    
    Dim formulaSeparator As String
    
    'Determine the formula separator for the current language settings
    formulaSeparator = Application.International(xlListSeparator)
    
    'Add the IFERROR function to the formula
    AddIfError = "=IFERROR(" & Right(formula, Len(formula) - Len(formulaSeparator)) & formulaSeparator & errorValue & ")"
End Function


