Attribute VB_Name = "domain"
Option Explicit
Dim ab_domain As String

Function adding_domain_return(add_domain, domain_return_index)
Select Case domain_return_index
Case 0
ab_domain = "i"
Case 1
ab_domain = "e"
End Select
add_domain = Chr(38) & "as_dt=" & ab_domain
End Function

