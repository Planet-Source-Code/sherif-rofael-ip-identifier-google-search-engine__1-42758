Attribute VB_Name = "numb_of_res"
Option Explicit
Dim add_res As String
Dim res As Integer

Function results_per_page(add_res, index_res)
res = 0
Select Case index_res
Case 0
res = 10
Case 1
res = 20
Case 2
res = 50
Case 3
res = 100
Case 4
res = 200
End Select
If res = 0 Then res = 10
add_res = Chr(38) & "num=" & res
End Function
