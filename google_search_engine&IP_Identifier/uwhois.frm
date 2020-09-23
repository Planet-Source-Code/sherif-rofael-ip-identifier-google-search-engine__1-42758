VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1950
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10005
   Icon            =   "uwhois.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4440
      Top             =   720
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox query 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"uwhois.frx":0E42
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image5 
      Height          =   510
      Left            =   6360
      Picture         =   "uwhois.frx":0F24
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3540
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   0
      Picture         =   "uwhois.frx":11A4
      Top             =   1440
      Width           =   2145
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   2040
      Picture         =   "uwhois.frx":1AAF
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5640
      MouseIcon       =   "uwhois.frx":1CF7
      MousePointer    =   99  'Custom
      Picture         =   "uwhois.frx":2001
      Top             =   1440
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "uwhois.frx":23DF
      Top             =   0
      Width           =   4830
   End
   Begin VB.Menu file_domain 
      Caption         =   "File"
      Begin VB.Menu new_domain 
         Caption         =   "New"
      End
      Begin VB.Menu gog_sea 
         Caption         =   "Search Google"
      End
      Begin VB.Menu exit_domain 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu vote_domain 
      Caption         =   "Vote for the program"
      Begin VB.Menu sitee 
         Caption         =   "@ my site"
      End
      Begin VB.Menu pscoddee 
         Caption         =   "@ planet source code"
      End
   End
   Begin VB.Menu aabb 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim domain_url As String
Dim IPP As String
Dim computer As String
Dim IntStatus As String
Dim text_to_scrool As String
Dim counter1 As Integer
Dim counter As Integer
Dim scrooling As String
Dim msg As String
Dim pscode_site As String
Dim mysite As String





Private Sub aabb_Click()
msg = MsgBox(" This is 'Google Search Engine' program designed by 'sherif rofael' mailto:ya3amo@hotmail.com ", vbInformation, "About me")
End Sub

Private Sub exit_domain_Click()
End
End Sub

Private Sub Form_Load()
Call getipp(IPP)
Call get_form_caption
End Sub

Private Sub gog_sea_Click()
Unload Me
Form1.Visible = True
End Sub

Private Sub Image2_Click()
domain_url = "http://www.uwhois.com/cgi/whois.cgi?User=Default&query=" & query.Text
Call RunBrowser(domain_url, 10, 1)
End Sub


Public Function getipp(IPP)
On Error GoTo skipping:
IPP = Winsock1.LocalIP
skipping:
End Function


Public Sub get_form_caption()
Call get_computer_name(computer)
Call InternetStatus(IntStatus)
text_to_scrool = "HI, " & computer & " & Welcome to the program 'Domain & IP Holder Identifier' , Designed by 'Sherif Rofael' in 27-th jan. 2003 mailto:ya3amo@hotmail.com , Thanks " & computer & " for using the program ," & IntStatus
For counter1 = 1 To Len(text_to_scrool) / 3
text_to_scrool = Chr(32) & text_to_scrool
Next counter1
End Sub

Public Function InternetStatus(IntStatus)
If InternetCheckConnection("http://www.pscode.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
       IntStatus = "No internet connection detected..."
    Else
        IntStatus = "You are connected to the Internet with an IP Address " & IPP
    End If
End Function

Private Sub new_domain_Click()
Unload Me
Form2.Visible = True
End Sub

Private Sub pscoddee_Click()
pscode_site = "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42744&lngWId=1"
Call RunBrowser(pscode_site, 10, 1)
End Sub

Private Sub sitee_Click()
mysite = "http://vbsherif.members.easyspace.com/google/vote.htm"
Call RunBrowser(mysite, 10, 1)
End Sub

Private Sub Timer1_Timer()
counter = counter + 1
If counter > Len(text_to_scrool) Then
counter = 0
End If
scrooling = Right(text_to_scrool, Len(text_to_scrool) - counter)
Form2.Caption = scrooling
End Sub

Function get_computer_name(computer)
Dim compname As String * 256
Call GetComputerName(compname, 256)
computer = Left(compname, InStr(compname, Chr(0)) - 1)
End Function
