Attribute VB_Name = "file_type"
Option Explicit
Dim ab_return As String

Function adding_return(add_return, return_index)
Select Case return_index
Case 0
ab_return = "i"
Case 1
ab_return = "e"
End Select
add_return = Chr(38) & "as_ft=" & ab_return
End Function
