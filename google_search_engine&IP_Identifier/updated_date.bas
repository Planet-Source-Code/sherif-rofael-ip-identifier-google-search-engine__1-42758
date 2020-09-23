Attribute VB_Name = "updated_date"
Option Explicit
Dim ab_update As String

Function updatedd_date(update_index, add_update)
Select Case update_index
Case 0
ab_update = ""
Case 1
ab_update = "m1"
Case 2
ab_update = "m2"
Case 3
ab_update = "m3"
Case 4
ab_update = "m6"
Case 5
ab_update = "m9"
Case 6
ab_update = "y"
End Select
add_update = Chr(38) & "as_qdr=" & ab_update
End Function
