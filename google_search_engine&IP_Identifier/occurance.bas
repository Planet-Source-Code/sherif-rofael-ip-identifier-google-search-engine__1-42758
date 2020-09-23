Attribute VB_Name = "occurance"
Option Explicit
Dim occ_ab As String


Function occur(occ_index, add_occur)
Select Case occ_index
Case 0
occ_ab = "any"
Case 1
occ_ab = "title"
Case 2
occ_ab = "body"
Case 3
occ_ab = "url"
Case 4
occ_ab = "links"
End Select
add_occur = Chr(38) & "as_occt=" & occ_ab
End Function
