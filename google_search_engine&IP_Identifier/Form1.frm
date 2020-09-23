VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   285
   ClientTop       =   840
   ClientWidth     =   6375
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0E42
   ScaleHeight     =   6855
   ScaleWidth      =   6375
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5520
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   28
      Top             =   5640
      Width           =   5415
   End
   Begin VB.OptionButton filter 
      Caption         =   "FILTER USING SAFE SEARCH"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   5280
      Width           =   2655
   End
   Begin VB.OptionButton no_filter 
      Caption         =   "NO FILTERING "
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   5040
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.ComboBox last_update_is 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox occurance_location 
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox site_search 
      Height          =   405
      Left            =   3960
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.ListBox domain_return 
      Height          =   450
      ItemData        =   "Form1.frx":114C
      Left            =   120
      List            =   "Form1.frx":1156
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox file_format 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Text            =   "FORMAT TYPE"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.ListBox return_or_not 
      Height          =   450
      ItemData        =   "Form1.frx":1167
      Left            =   120
      List            =   "Form1.frx":1171
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox search_without 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox at_least 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox exact_phrase 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox search_for 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GOOGLE SEARCH"
      Height          =   375
      Left            =   1320
      MouseIcon       =   "Form1.frx":1182
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6360
      Width           =   3135
   End
   Begin VB.ComboBox number_of_results 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Text            =   "Number of Results Per Page"
      Top             =   5040
      Width           =   2655
   End
   Begin VB.ComboBox only_in_lang 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "ALL LANGUAGES"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox disp_lang 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Text            =   "CHOOSE LANGUAGE"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "SEARCH ENGINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   29
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":148C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "e.g. google.com, .org  "
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Return web pages updated in the :"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "(english by default)"
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Return results where my terms occur:"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   "return results from the site or domain"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "return results of the file format:"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "DISPLAY RESULTS IN:"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Return pages written in"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "FIND RESULTS without the words:"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "FIND RESULTS with at least one of the words:"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "FIND RESULTS with the exact phrase:"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "FIND RESULTS with all of the words:"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   3375
   End
   Begin VB.Menu filee 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu sea_dom 
         Caption         =   "Search For a Domain"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu voting 
      Caption         =   "Vote for the program"
      Begin VB.Menu site 
         Caption         =   "@ my site"
         Shortcut        =   ^M
      End
      Begin VB.Menu pscode 
         Caption         =   "@planet source code"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu about_me 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim index_res As Integer
Dim add_res As String
Dim ind_dis As String
Dim add_disp As String
Dim pages_in As String
Dim add_search_lang As String
Dim return_index As Integer
Dim add_return As String
Dim search_for_what As String
Dim search_with_exact As String
Dim search_with_at_least As String
Dim without_these_words As String
Dim ab_index As Integer
Dim add_format As String
Dim domain_return_index As Integer
Dim add_domain As String
Dim add_search_site As String
Dim occ_index As Integer
Dim add_occur As String
Dim safe_s As String
Dim update_index As Integer
Dim add_update As String
Dim url As String
Dim search_for_string As String
Dim exact_phrase_string As String
Dim at_least_string As String
Dim search_without_string As String
Dim computer As String
Dim text_to_scrool As String
Dim counter As Integer
Dim counter1 As Integer
Dim scrooling As String
Dim msg As Boolean
Dim pscode_site As String
Dim mysite As String
Public IPaddress As String
Dim IntStatus As String




Function filtering()
If no_filter.Value = True Then safe_s = Chr(38) & "safe=images"
If filter.Value = True Then safe_s = Chr(38) & "safe=active"
End Function

Function getstring()
Call get_search_text(search_for.Text, search_for_string)
Call get_search_text(exact_phrase.Text, exact_phrase_string)
Call get_search_text(at_least.Text, at_least_string)
Call get_search_text(search_without.Text, search_without_string)
End Function

Function searching()
Call getstring
search_for_what = Chr(38) & "as_q=" & search_for_string 'search for all these words
search_with_exact = Chr(38) & "as_epq=" & exact_phrase_string 'exact prase
search_with_at_least = Chr(38) & "as_oq=" & at_least_string 'with @ least one of these
without_these_words = Chr(38) & "as_eq=" & search_without_string 'with @ least one of these
add_search_site = Chr(38) & "as_sitesearch=" & site_search
End Function

Private Sub about_me_Click()
msg = MsgBox(" This is 'Google Search Engine' program designed by 'sherif rofael' mailto:ya3amo@hotmail.com ", vbInformation, "About me")

End Sub

Private Sub at_least_Change()
Call execute

End Sub

Private Sub Command1_Click()
Call execute
Call RunBrowser(url, 10, 1)
End Sub

Private Sub disp_lang_Click()
Call execute
End Sub

Private Sub domain_return_Click()
Call execute
End Sub

Private Sub exact_phrase_Change()
Call execute
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub file_format_Click()
Call execute
End Sub

Private Sub filter_Click()
Call execute
End Sub

Private Sub Form_Load()
Call getip(IPaddress)
Call get_form_caption
Call add_to_disp_lang
Call only_pages
Call eresults_per_page
Call formatt_type
Call occurance_loc
occurance_location.SelText = occurance_location.List(0)
Call date_of_last
last_update_is.SelText = last_update_is.List(0)
End Sub



Public Function add_to_disp_lang()
Call disp_lang.AddItem("Arabic", 0)
Call disp_lang.AddItem("Bulgarian", 1)
Call disp_lang.AddItem("Catalan", 2)
Call disp_lang.AddItem("Chinese Simplified", 3)
Call disp_lang.AddItem("Chinese Traditional", 4)
Call disp_lang.AddItem("Croatian", 5)
Call disp_lang.AddItem("Czech", 6)
Call disp_lang.AddItem("Danish", 7)
Call disp_lang.AddItem("Dutch", 8)
Call disp_lang.AddItem("English", 9)
Call disp_lang.AddItem("Estonian", 10)
Call disp_lang.AddItem("Finnish", 11)
Call disp_lang.AddItem("French", 12)
Call disp_lang.AddItem("German", 13)
Call disp_lang.AddItem("Greek", 14)
Call disp_lang.AddItem("Hebrew", 15)
Call disp_lang.AddItem("Hungarian", 16)
Call disp_lang.AddItem("Icelandic", 17)
Call disp_lang.AddItem("Indonesian", 18)
Call disp_lang.AddItem("Italian", 19)
Call disp_lang.AddItem("Japanese", 20)
Call disp_lang.AddItem("Korean", 21)
Call disp_lang.AddItem("Latvian", 22)
Call disp_lang.AddItem("Lithuanian", 23)
Call disp_lang.AddItem("Norwegian", 24)
Call disp_lang.AddItem("Polish", 25)
Call disp_lang.AddItem("Portuguese", 26)
Call disp_lang.AddItem("Romanian", 27)
Call disp_lang.AddItem("Russian", 28)
Call disp_lang.AddItem("Serbian", 29)
Call disp_lang.AddItem("Slovak", 30)
Call disp_lang.AddItem("Slovenian", 31)
Call disp_lang.AddItem("Spanish", 32)
Call disp_lang.AddItem("Swedish", 33)
Call disp_lang.AddItem("Turkish", 34)
End Function

Public Function only_pages()
Call only_in_lang.AddItem("Arabic", 0)
Call only_in_lang.AddItem("Bulgarian", 1)
Call only_in_lang.AddItem("Catalan", 2)
Call only_in_lang.AddItem("Chinese Simplified", 3)
Call only_in_lang.AddItem("Chinese Traditional", 4)
Call only_in_lang.AddItem("Croatian", 5)
Call only_in_lang.AddItem("Czech", 6)
Call only_in_lang.AddItem("Danish", 7)
Call only_in_lang.AddItem("Dutch", 8)
Call only_in_lang.AddItem("English", 9)
Call only_in_lang.AddItem("Estonian", 10)
Call only_in_lang.AddItem("Finnish", 11)
Call only_in_lang.AddItem("French", 12)
Call only_in_lang.AddItem("German", 13)
Call only_in_lang.AddItem("Greek", 14)
Call only_in_lang.AddItem("Hebrew", 15)
Call only_in_lang.AddItem("Hungarian", 16)
Call only_in_lang.AddItem("Icelandic", 17)
Call only_in_lang.AddItem("Indonesian", 18)
Call only_in_lang.AddItem("Italian", 19)
Call only_in_lang.AddItem("Japanese", 20)
Call only_in_lang.AddItem("Korean", 21)
Call only_in_lang.AddItem("Latvian", 22)
Call only_in_lang.AddItem("Lithuanian", 23)
Call only_in_lang.AddItem("Norwegian", 24)
Call only_in_lang.AddItem("Polish", 25)
Call only_in_lang.AddItem("Portuguese", 26)
Call only_in_lang.AddItem("Romanian", 27)
Call only_in_lang.AddItem("Russian", 28)
Call only_in_lang.AddItem("Serbian", 29)
Call only_in_lang.AddItem("Slovak", 30)
Call only_in_lang.AddItem("Slovenian", 31)
Call only_in_lang.AddItem("Spanish", 32)
Call only_in_lang.AddItem("Swedish", 33)
Call only_in_lang.AddItem("Turkish", 34)
Call only_in_lang.AddItem("any language", 35)
End Function
Public Function occurance_loc()
Call occurance_location.AddItem("anywhere in the page", 0)
Call occurance_location.AddItem("in the title of the page", 1)
Call occurance_location.AddItem("in the text of the page", 2)
Call occurance_location.AddItem("in the url of the page", 3)
Call occurance_location.AddItem("in links to the page", 4)
End Function

Public Function formatt_type()
Call file_format.AddItem("any format", 0)
Call file_format.AddItem("Adobe Acrobat PDF (.pdf)", 1)
Call file_format.AddItem("Adobe Postscript (.ps)", 2)
Call file_format.AddItem("Microsoft Word (.doc)", 3)
Call file_format.AddItem("Microsoft Excel (.xls)", 4)
Call file_format.AddItem("Microsoft Powerpoint (.ppt)", 5)
Call file_format.AddItem("Rich Text Format (.rtf)", 6)
End Function

Private Sub eresults_per_page()
Call number_of_results.AddItem("10 RESULTS", 0)
Call number_of_results.AddItem("20 RESULTS", 1)
Call number_of_results.AddItem("50 RESULTS", 2)
Call number_of_results.AddItem("100 RESULTS", 3)
Call number_of_results.AddItem("200 RESULTS", 4)
End Sub


Private Sub date_of_last()
Call last_update_is.AddItem("any time", 0)
Call last_update_is.AddItem("past 1 month ", 1)
Call last_update_is.AddItem("past 2 months", 2)
Call last_update_is.AddItem("past 3 months", 3)
Call last_update_is.AddItem("past 6 months", 4)
Call last_update_is.AddItem("past 9 months", 5)
Call last_update_is.AddItem("past year", 6)
End Sub

Function execute()
index_res = number_of_results.ListIndex
Call results_per_page(add_res, index_res)

ind_dis = disp_lang.ListIndex
Call adding_display_language(add_disp, ind_dis)

pages_in = only_in_lang.ListIndex
Call adding_searching_language(add_search_lang, pages_in)

Call searching

return_index = return_or_not.ListIndex
Call adding_return(add_return, return_index)

domain_return_index = domain_return.ListIndex
Call adding_domain_return(add_domain, domain_return_index)

ab_index = file_format.ListIndex
Call format_abb(ab_index, add_format)

occ_index = occurance_location.ListIndex
Call occur(occ_index, add_occur)

update_index = last_update_is.ListIndex
Call updatedd_date(update_index, add_update)

Call filtering

Call geturl
End Function

Function geturl()
url = "http://www.google.com/search?" & safe_s & search_for_what & search_with_exact & search_with_at_least & without_these_words & add_search_site & add_res & add_disp & add_search_lang & add_return & add_domain & add_format & add_occur & add_update
Text1.Text = url
'Call label_caption
'the_url.Caption = url
End Function

Function get_computer_name(computer)
Dim compname As String * 256
Call GetComputerName(compname, 256)
computer = Left(compname, InStr(compname, Chr(0)) - 1)
End Function

Public Sub get_form_caption()
Call get_computer_name(computer)
Call InternetStatus(IntStatus)
text_to_scrool = "HI, " & computer & " & Welcome to the program 'Google Search Engine' , Designed by 'Sherif Rofael' in 25-th jan. 2003 mailto:ya3amo@hotmail.com , Thanks " & computer & " for using the program ," & IntStatus
For counter1 = 1 To Len(text_to_scrool) / 3
text_to_scrool = Chr(32) & text_to_scrool
Next counter1
End Sub

Private Sub last_update_is_Click()
Call execute
End Sub

Private Sub new_Click()
Unload Me
Form1.Visible = True
End Sub

Private Sub no_filter_Click()
Call execute
End Sub

Private Sub number_of_results_Click()
Call execute
End Sub

Private Sub occurance_location_Click()
Call execute
End Sub



Private Sub only_in_lang_Click()
Call execute
End Sub

Private Sub pscode_Click()
pscode_site = "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42744&lngWId=1"
Call RunBrowser(pscode_site, 10, 1)
End Sub

Private Sub return_or_not_Click()
Call execute
End Sub

Private Sub sea_dom_Click()
Unload Me
Form2.Visible = True
End Sub

Private Sub search_for_Change()
Call execute

End Sub

Private Sub search_without_Change()
Call execute
End Sub

Private Sub site_Click()
mysite = "http://vbsherif.members.easyspace.com/google/vote.htm"
Call RunBrowser(mysite, 10, 1)
End Sub

Private Sub site_search_change()
Call execute
End Sub



Private Sub Timer1_Timer()
counter = counter + 1
If counter > Len(text_to_scrool) Then
counter = 0
End If
scrooling = Right(text_to_scrool, Len(text_to_scrool) - counter)
Form1.Caption = scrooling
End Sub

Function label_caption()
Dim l, r, v, icounter, numm, y
Text1.Text = ""
numm = 60
For icounter = 1 To Int(Len(url) / numm) + 1
l = Left(url, numm * icounter)
r = Right(l, numm)
'v = v & r & Chr(32)
Text1.Text = Text1.Text & r & vbCrLf
Next icounter
'y = ((Len(url) / numm) - Int(Len(url) / numm)) * numm
'r = Right(l, y)
'Text1.Text = Text1.Text & r & vbCrLf
'Label14.Caption = v
End Function


Public Function getip(IPaddress)
On Error GoTo skipping:
IPaddress = Winsock1.LocalIP
skipping:
End Function

Public Function InternetStatus(IntStatus)
If InternetCheckConnection("http://www.pscode.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
       IntStatus = "No internet connection detected..."
    Else
        IntStatus = "You are connected to the Internet with an IP Address " & IPaddress
    End If
End Function

