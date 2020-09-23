Attribute VB_Name = "abbreviation"
Option Explicit
Dim add_search_lang  As String
Dim add_disp  As String
Dim ab_display  As String
Dim ab_in  As String

Function language(index_lang, ab)
Select Case index_lang
Case 0
ab = "ar"
Case 1
ab = "bg"
Case 2
ab = "ca"
Case 3
ab = "zh-CN"
Case 4
ab = "zh-TW"
Case 5
ab = "hr"
Case 6
ab = "cs"
Case 7
ab = "da"
Case 8
ab = "nl"
Case 9
ab = "en"
Case 10
ab = "et"
Case 11
ab = "fi"
Case 12
ab = "fr"
Case 13
ab = "de"
Case 14
ab = "el"
Case 15
ab = "iw"
Case 16
ab = "hu"
Case 17
ab = "is"
Case 18
ab = "id"
Case 19
ab = "it"
Case 20
ab = "ja"
Case 21
ab = "ko"
Case 22
ab = "lv"
Case 23
ab = "lt"
Case 24
ab = "no"
Case 25
ab = "pl"
Case 26
ab = "pt"
Case 27
ab = "ro"
Case 28
ab = "ru"
Case 29
ab = "sr"
Case 30
ab = "sk"
Case 31
ab = "sl"
Case 32
ab = "es"
Case 33
ab = "sv"
Case 34
ab = "tr"
End Select
End Function



Function adding_display_language(add_disp, ind_dis)
Call language(ind_dis, ab_display)
add_disp = Chr(38) & "hl=" & ab_display
End Function

Function adding_searching_language(add_search_lang, pages_in)
Call language(pages_in, ab_in)
add_search_lang = Chr(38) & "lr=lang_" & ab_in
End Function
