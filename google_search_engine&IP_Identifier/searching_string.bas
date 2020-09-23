Attribute VB_Name = "searching_string"
Option Explicit
Dim word As String
Dim ml As String
Dim mr As String
Dim i As Integer

Function get_search_text(ipexp, word)
word = ""
For i = 1 To Len(ipexp)
ml = Left(ipexp, i)
mr = Right(ml, 1)
If mr = Chr(32) Then
word = word & "+"
Else
word = word & mr
End If
Next i
End Function
