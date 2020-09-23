Attribute VB_Name = "format_abbreviation"
Option Explicit
Dim ab_format As String

Function format_abb(ab_index, add_format)
Select Case ab_index
Case 0
ab_format = ""
Case 1
ab_format = "pdf"
Case 2
ab_format = "ps"
Case 3
ab_format = "doc"
Case 4
ab_format = "xls"
Case 5
ab_format = "ppt"
Case 6
ab_format = "rtf"
End Select
add_format = Chr(38) & "as_filetype=" & ab_format
End Function

