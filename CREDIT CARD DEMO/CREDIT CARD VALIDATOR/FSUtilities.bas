Attribute VB_Name = "FSUtilities"
Option Explicit

'==============================================================================
' YOU ARE FREE TO USE THIS CODE IN YOUR OWN VB PROJECTS PROVIDED NO
' CHANGES ARE MADE TO THE ORIGINAL SOURCE CODE. PLEASE REPORT BUGS
' TO const71@yahoo.com
' Copyright(c) 2003 Constantin Nterekas
' http://www.foundationssoftware.com
'==============================================================================

Public Const ALPHAS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const NUMBERS = "0123456789"

'==============================================================================
'      METHOD: IsAlpha
' DESCRIPTION: Is the specified value an alphabetic character?
'==============================================================================
Public Function IsAlpha(ByVal value As String) As Boolean
    IsAlpha = IsInFilter(value, ALPHAS)
End Function
'==============================================================================
'      METHOD: IsNumber
' DESCRIPTION: Is the specified value a number?
'==============================================================================
Public Function IsNumber(ByVal value As String) As Boolean
    IsNumber = IsInFilter(value, NUMBERS)
End Function
'==============================================================================
'      METHOD: IsAlphaNumeric
' DESCRIPTION: Is the specified value an alphanumeric character?
'==============================================================================
Public Function IsAlphaNumeric(ByVal value As String) As Boolean
    IsAlphaNumeric = IsInFilter(value, ALPHAS & NUMBERS)
End Function
'==============================================================================
'      METHOD: IsInFilter
' DESCRIPTION: Each character in sMyString is checked against the filter set.
'              If any character is not found in the filter set, false is
'              returned, otherwise true.
'==============================================================================
Public Function IsInFilter(ByVal sMyString As String, ByVal filter As String) As Boolean
    Dim i As Long
    IsInFilter = True
    For i = 1 To Len(sMyString)
        If InStr(1, filter, Mid(sMyString, i, 1), vbTextCompare) = 0 Then
            IsInFilter = False
            Exit For
        End If
    Next i
End Function
