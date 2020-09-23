Attribute VB_Name = "computer_name"
Option Explicit

'computer name
'***********************************
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'***********************************

