VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5880
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.OptionButton uwhois2 
      Caption         =   "Identify the registered holder of a domain name"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.OptionButton google 
      Caption         =   "Search Google (Advanced Search)"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.OptionButton uwhois1 
      Caption         =   "Identify the registered holder of an IP Adress"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   4320
      Picture         =   "Form3.frx":0E42
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   3600
      Picture         =   "Form3.frx":22F2
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1740
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim domain_url As String
Dim IPP As String
Dim computer
Dim IntStatus
Dim text_to_scrool
Dim counter1
Dim counter
Dim scrooling

Private Sub Command1_Click()
If uwhois1.Value = True Then Unload Me: Form2.Visible = True
If uwhois2.Value = True Then Unload Me: Form2.Visible = True
If google.Value = True Then Unload Me: Form1.Visible = True


End Sub

Private Sub Form_Load()
Call get_form_caption
End Sub

Public Sub get_form_caption()
Call get_computer_name(computer)
text_to_scrool = "HI, " & computer & " & Welcome to the program 'Google Search Engine & Domain & IP Holder Identifier' , Designed by 'Sherif Rofael' in 27-th jan. 2003 mailto:ya3amo@hotmail.com , Thanks " & computer & " for using the program ,"
For counter1 = 1 To Len(text_to_scrool) / 3
text_to_scrool = Chr(32) & text_to_scrool
Next counter1
End Sub









Private Sub Timer1_Timer()
counter = counter + 1
If counter > Len(text_to_scrool) Then
counter = 0
End If
scrooling = Right(text_to_scrool, Len(text_to_scrool) - counter)
Form3.Caption = scrooling
End Sub

Function get_computer_name(computer)
Dim compname As String * 256
Call GetComputerName(compname, 256)
computer = Left(compname, InStr(compname, Chr(0)) - 1)
End Function

