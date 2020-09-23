Attribute VB_Name = "Internet_connection"
Option Explicit

Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long


 
