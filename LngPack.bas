Attribute VB_Name = "LngPack"
'Language Pack
'© Scythe Scythe@cablenet.de

'This is nearly the same version i released on psc

'Funtions:
'---------

'SaveLanguage
'Creates a languagefile for ur APP

'LoadLanguage
'Loads a Languagefile and chages the language

'GetLanguageID
'Get the windows Language ID

'GetLanguageNameLong
'Gets the Language Name incl. Loacl info (English (New Zealand))

'GetLanguageName
'Gets the Language Name (English)


Option Explicit


'Find the Windows Language
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
Private Const LOCALE_ILANGUAGE = &H1


Public Sub SaveLanguage(Frm As Form, LanguageFile As String)
 Dim Contrl As Control
 Dim CTyp As String
 Dim I As Long
 Dim Tmp As String

 On Error GoTo ErrOut
 Open LanguageFile For Output As #1

  For Each Contrl In Frm.Controls

  'Get the Controltype
  CTyp = TypeName(Contrl)
  Select Case (CTyp)
   'Case "Label", "Frame", "CheckBox", "CommandButton", "OptionButton"

   'We can handle Contrl like the control
   'Now save the data we need
  Case "ComboBox", "ListBox"
   Print #1, "#" & Contrl.ListCount
   Print #1, Contrl.Name
   Print #1, CheckForIndex(Contrl)
   For I = 0 To Contrl.ListCount
    Print #1, Contrl.List(I)
   Next I

  Case "TabStrip"
   Print #1, "~" & Contrl.Tabs.Count
   Print #1, Contrl.Name
   Print #1, CheckForIndex(Contrl)
   For I = 1 To Contrl.Tabs.Count
    Print #1, Contrl.Tabs(I).Caption
    Print #1, Contrl.Tabs(I).ToolTipText
   Next I

  Case "ListView"
   Print #1, "+" & Contrl.ColumnHeaders.Count
   Print #1, Contrl.Name
   Print #1, CheckForIndex(Contrl)
   For I = 1 To Contrl.ColumnHeaders.Count
    Print #1, Contrl.ColumnHeaders(I).Text
   Next I

  Case Else
   'Controls With Caption & Tooltip
   Tmp = TestCtrl(Contrl)
   If Tmp <> "" Then
    Print #1, Tmp
   End If

  End Select
 Next
Close
Exit Sub
ErrOut:
MsgBox "Error saving Language" & vbCrLf & Err.Number, vbCritical, "Save Language"
Close
End Sub

'Load Language
'The Heart of this module
'If u want it short
'Only distribute this function with ur app
Public Function LoadLanguage(Frm As Form, ByVal LanguageFile As String, Optional ShowErrors As Boolean = True) As Long
 Dim I As Long
 Dim Ctr As Long
 Dim Tmp As String
 Dim Tmp1 As String
 Dim CtrlName As String
 Dim CtrlIndex As Long

 'We can access controls by Name
 'like Control("Label1).Caption = "Label number one"

 On Error GoTo ErrOut

 'Open the LanguageFile
 Open LanguageFile For Input As #1
  Do Until EOF(1)

   Line Input #1, Tmp       'Get the Header

   If Len(Tmp) > 1 Then     'Extract the counter if there is one
    Ctr = Val(Right$(Tmp, Len(Tmp) - 1))
   End If

   Line Input #1, CtrlName  'Get ControlName

   Line Input #1, Tmp1      'Get Index
   CtrlIndex = Val(Tmp1)

   Select Case Left$(Tmp, 1)


   Case "#" 'Combobox
    For I = 0 To Ctr 'Now read all lines for the Combobox
     Line Input #1, Tmp
     If CtrlIndex = -1 Then '-1 = No Index
      Frm.Controls(CtrlName).List(I) = Tmp
     Else                   'Index Control [Combo1(3)]
      Frm.Controls(CtrlName)(CtrlIndex).List(I) = Tmp
     End If
    Next I

   Case "~" 'Tabstrip
    For I = 1 To Ctr
     Line Input #1, Tmp
     Line Input #1, Tmp1
     If CtrlIndex = -1 Then '-1 = No Index
      Frm.Controls(CtrlName).Tabs(I).Caption = Tmp
      Frm.Controls(CtrlName).Tabs(I).ToolTipText = Tmp1
     Else
      Frm.Controls(CtrlName)(CtrlIndex).Tabs(I).Caption = Tmp
      Frm.Controls(CtrlName)(CtrlIndex).Tabs(I).ToolTipText = Tmp1
     End If
    Next I

   Case "+" 'ListView
    For I = 1 To Ctr
     Line Input #1, Tmp
     If CtrlIndex = -1 Then
      Frm.Controls(CtrlName).ColumnHeaders(I).Text = Tmp
     Else
      Frm.Controls(CtrlName)(CtrlIndex).ColumnHeaders(I).Text = Tmp
     End If
    Next I

   Case "^" 'Control with Only Caption
    Line Input #1, Tmp
    If CtrlIndex = -1 Then
     Frm.Controls(CtrlName).Caption = Tmp
    Else
     Frm.Controls(CtrlName)(CtrlIndex).Caption = Tmp
    End If

   Case "°" 'Control with Only Tooltip
    Line Input #1, Tmp
    If CtrlIndex = -1 Then
     Frm.Controls(CtrlName).ToolTipText = Tmp
    Else
     Frm.Controls(CtrlName)(CtrlIndex).ToolTipText = Tmp
    End If

   Case "*" 'Control with Caption & ToolTipText
    Line Input #1, Tmp
    Line Input #1, Tmp1
    If CtrlIndex = -1 Then
     Frm.Controls(CtrlName).Caption = Tmp
     Frm.Controls(CtrlName).ToolTipText = Tmp1
    Else
     Frm.Controls(CtrlName)(CtrlIndex).Caption = Tmp
     Frm.Controls(CtrlName)(CtrlIndex).ToolTipText = Tmp1
    End If
   End Select
  Loop
 Close
 Exit Function
ErrOut:
 If ShowErrors Then
  If Err.Number = 53 Then
   MsgBox "Cant find Language Pack", vbCritical, "Load Language"
   ElseIf Err.Number <> 62 Then
   MsgBox "Error loading Language File" & vbCrLf & Err.Description, vbCritical, "Load Language"
  End If
 End If
 LoadLanguage = Err.Number
 Close
End Function



'See if the Control has an Index
Private Function CheckForIndex(Ctrl As Control) As String
 On Error GoTo ErrOut
 CheckForIndex = Str(Ctrl.Index)
 Exit Function
ErrOut:
 CheckForIndex = "-1"
End Function

'Check for Captions / Tooltips
Private Function TestCtrl(Ctrl As Control) As String
 Dim Tmp As String
 Dim Tmp2 As String
 Dim X As Long
 Tmp = Ctrl.Name & vbCrLf
 Tmp = Tmp & CheckForIndex(Ctrl)

 Tmp2 = GetCaption(Ctrl)
 If Tmp2 <> "" Then
  Tmp = Tmp & vbCrLf & Tmp2
  X = X + 1
 End If

 Tmp2 = GetTooltip(Ctrl)
 If Tmp2 <> "" Then
  Tmp = Tmp & vbCrLf & Tmp2
  X = X + 2
 End If

 Select Case X
 Case 1 'Only Caption
  TestCtrl = "^" & vbCrLf & Tmp
 Case 2 'Only Tooltip
  TestCtrl = "°" & vbCrLf & Tmp
 Case 3 'Caption & Tooltip
  TestCtrl = "*" & vbCrLf & Tmp
 End Select

End Function

'Get the Caption
Private Function GetCaption(Ctrl As Control) As String
 On Error GoTo ErrOut
 GetCaption = Ctrl.Caption
ErrOut:
End Function

'Get the Tooltip
Private Function GetTooltip(Ctrl As Control) As String
 On Error GoTo ErrOut
 GetTooltip = Ctrl.ToolTipText
ErrOut:
End Function



Public Function GetLanguageID() As Long
 Dim Buf As String
 Dim Res As Long
 Dim Lgh As String
 '# Get the Language
 'Get the lenght of the LanguageID
 Lgh = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_ILANGUAGE, Buf, 0) - 1
 'Fill the Buffer with spaces
 Buf = Space(Lgh + 1)
 'Get the Language ID
 Res = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_ILANGUAGE, Buf, Lgh)
 'Cut the String
 GetLanguageID = Val(Left$("&H" & Buf, Lgh + 2))
End Function

'To easy to write any word about it :o)
Public Function GetLanguageNameLong(LangId As Long) As String
 Select Case LangId
 Case 1078
  GetLanguageNameLong = "Afrikaans"
 Case 1052
  GetLanguageNameLong = "Albanian"
 Case 5121
  GetLanguageNameLong = "Arabic (Algeria)"
 Case 15361
  GetLanguageNameLong = "Arabic (Bahrain)"
 Case 3073
  GetLanguageNameLong = "Arabic (Egypt)"
 Case 2049
  GetLanguageNameLong = "Arabic (Iraq)"
 Case 11265
  GetLanguageNameLong = "Arabic (Jordan)"
 Case 13313
  GetLanguageNameLong = "Arabic (Kuwait)"
 Case 12289
  GetLanguageNameLong = "Arabic (Lebanon)"
 Case 4097
  GetLanguageNameLong = "Arabic (Libya)"
 Case 6145
  GetLanguageNameLong = "Arabic (Morocco)"
 Case 8193
  GetLanguageNameLong = "Arabic (Oman)"
 Case 16385
  GetLanguageNameLong = "Arabic (Qatar)"
 Case 1025
  GetLanguageNameLong = "Arabic (Saudi Arabia)"
 Case 10241
  GetLanguageNameLong = "Arabic (Syria)"
 Case 7169
  GetLanguageNameLong = "Arabic (Tunisia)"
 Case 14337
  GetLanguageNameLong = "Arabic (U.A.E.)"
 Case 9217
  GetLanguageNameLong = "Arabic (Yemen)"
 Case 1067
  GetLanguageNameLong = "Armenian"
 Case 2092
  GetLanguageNameLong = "Azeri (Cyrillic)"
 Case 1068
  GetLanguageNameLong = "Azeri (Latin)"
 Case 1069
  GetLanguageNameLong = "Basque"
 Case 1059
  GetLanguageNameLong = "Belarusian"
 Case 1026
  GetLanguageNameLong = "Bulgarian"
 Case 1027
  GetLanguageNameLong = "Catalan"
 Case 3076
  GetLanguageNameLong = "Chinese (Hong Kong S.A.R.)"
 Case 5124
  GetLanguageNameLong = "Chinese (Macau S.A.R.)"
 Case 2052
  GetLanguageNameLong = "Chinese (PRC)"
 Case 4100
  GetLanguageNameLong = "Chinese (Singapore)"
 Case 1028
  GetLanguageNameLong = "Chinese (Taiwan)"
 Case 1050
  GetLanguageNameLong = "Croatian"
 Case 1029
  GetLanguageNameLong = "Czech"
 Case 1030
  GetLanguageNameLong = "Danish"
 Case 1125
  GetLanguageNameLong = "Divehi"
 Case 2067
  GetLanguageNameLong = "Dutch (Belgium)"
 Case 1043
  GetLanguageNameLong = "Dutch (Netherlands)"
 Case 3081
  GetLanguageNameLong = "English (Australia)"
 Case 10249
  GetLanguageNameLong = "English (Belize)"
 Case 4105
  GetLanguageNameLong = "English (Canada)"
 Case 9225
  GetLanguageNameLong = "English (Caribbean)"
 Case 6153
  GetLanguageNameLong = "English (Ireland)"
 Case 8201
  GetLanguageNameLong = "English (Jamaica)"
 Case 5129
  GetLanguageNameLong = "English (New Zealand)"
 Case 13321
  GetLanguageNameLong = "English (Philippines)"
 Case 7177
  GetLanguageNameLong = "English (South Africa)"
 Case 11273
  GetLanguageNameLong = "English (Trinidad)"
 Case 2057
  GetLanguageNameLong = "English (United Kingdom)"
 Case 1033
  GetLanguageNameLong = "English (United States)"
 Case 12297
  GetLanguageNameLong = "English (Zimbabwe)"
 Case 1061
  GetLanguageNameLong = "Estonian"
 Case 1080
  GetLanguageNameLong = "Faroese"
 Case 1065
  GetLanguageNameLong = "Farsi"
 Case 1035
  GetLanguageNameLong = "Finnish"
 Case 2060
  GetLanguageNameLong = "French (Belgium)"
 Case 3084
  GetLanguageNameLong = "French (Canada)"
 Case 1036
  GetLanguageNameLong = "French (France)"
 Case 5132
  GetLanguageNameLong = "French (Luxembourg)"
 Case 6156
  GetLanguageNameLong = "French (Monaco)"
 Case 4108
  GetLanguageNameLong = "French (Switzerland)"
 Case 1071
  GetLanguageNameLong = "FYRO Macedonian"
 Case 1110
  GetLanguageNameLong = "Galician"
 Case 1079
  GetLanguageNameLong = "Georgian"
 Case 3079
  GetLanguageNameLong = "German (Austria)"
 Case 1031
  GetLanguageNameLong = "German (Germany)"
 Case 5127
  GetLanguageNameLong = "German (Liechtenstein)"
 Case 4103
  GetLanguageNameLong = "German (Luxembourg)"
 Case 2055
  GetLanguageNameLong = "German (Switzerland)"
 Case 1032
  GetLanguageNameLong = "Greek"
 Case 1095
  GetLanguageNameLong = "Gujarati"
 Case 1037
  GetLanguageNameLong = "Hebrew"
 Case 1081
  GetLanguageNameLong = "Hindi"
 Case 1038
  GetLanguageNameLong = "Hungarian"
 Case 1039
  GetLanguageNameLong = "Icelandic"
 Case 1057
  GetLanguageNameLong = "Indonesian"
 Case 1040
  GetLanguageNameLong = "Italian (Italy)"
 Case 2064
  GetLanguageNameLong = "Italian (Switzerland)"
 Case 1041
  GetLanguageNameLong = "Japanese"
 Case 1099
  GetLanguageNameLong = "Kannada"
 Case 1087
  GetLanguageNameLong = "Kazakh"
 Case 1111
  GetLanguageNameLong = "Konkani"
 Case 1042
  GetLanguageNameLong = "Korean"
 Case 1088
  GetLanguageNameLong = "Kyrgyz (Cyrillic)"
 Case 1062
  GetLanguageNameLong = "Latvian"
 Case 1063
  GetLanguageNameLong = "Lithuanian"
 Case 2110
  GetLanguageNameLong = "Malay (Brunei Darussalam)"
 Case 1086
  GetLanguageNameLong = "Malay (Malaysia)"
 Case 1102
  GetLanguageNameLong = "Marathi"
 Case 1104
  GetLanguageNameLong = "Mongolian (Cyrillic)"
 Case 1044
  GetLanguageNameLong = "Norwegian (Bokmal)"
 Case 2068
  GetLanguageNameLong = "Norwegian (Nynorsk)"
 Case 1045
  GetLanguageNameLong = "Polish"
 Case 1046
  GetLanguageNameLong = "Portuguese (Brazil)"
 Case 2070
  GetLanguageNameLong = "Portuguese (Portugal)"
 Case 1094
  GetLanguageNameLong = "Punjabi"
 Case 1048
  GetLanguageNameLong = "Romanian"
 Case 1049
  GetLanguageNameLong = "Russian"
 Case 1103
  GetLanguageNameLong = "Sanskrit"
 Case 3098
  GetLanguageNameLong = "Serbian (Cyrillic)"
 Case 2074
  GetLanguageNameLong = "Serbian (Latin)"
 Case 1051
  GetLanguageNameLong = "Slovak"
 Case 1060
  GetLanguageNameLong = "Slovenian"
 Case 11274
  GetLanguageNameLong = "Spanish (Argentina)"
 Case 16394
  GetLanguageNameLong = "Spanish (Bolivia)"
 Case 13322
  GetLanguageNameLong = "Spanish (Chile)"
 Case 9226
  GetLanguageNameLong = "Spanish (Colombia)"
 Case 5130
  GetLanguageNameLong = "Spanish (Costa Rica)"
 Case 7178
  GetLanguageNameLong = "Spanish (Dominican Republic)"
 Case 12298
  GetLanguageNameLong = "Spanish (Ecuador)"
 Case 17418
  GetLanguageNameLong = "Spanish (El Salvador)"
 Case 4106
  GetLanguageNameLong = "Spanish (Guatemala)"
 Case 18442
  GetLanguageNameLong = "Spanish (Honduras)"
 Case 3082
  GetLanguageNameLong = "Spanish (International Sort)"
 Case 2058
  GetLanguageNameLong = "Spanish (Mexico)"
 Case 19466
  GetLanguageNameLong = "Spanish (Nicaragua)"
 Case 6154
  GetLanguageNameLong = "Spanish (Panama)"
 Case 15370
  GetLanguageNameLong = "Spanish (Paraguay)"
 Case 10250
  GetLanguageNameLong = "Spanish (Peru)"
 Case 20490
  GetLanguageNameLong = "Spanish (Puerto Rico)"
 Case 1034
  GetLanguageNameLong = "Spanish (Traditional Sort)"
 Case 14346
  GetLanguageNameLong = "Spanish (Uruguay)"
 Case 8202
  GetLanguageNameLong = "Spanish (Venezuela)"
 Case 1089
  GetLanguageNameLong = "Swahili"
 Case 1053
  GetLanguageNameLong = "Swedish"
 Case 2077
  GetLanguageNameLong = "Swedish (Finland)"
 Case 1114
  GetLanguageNameLong = "Syriac"
 Case 1097
  GetLanguageNameLong = "Tamil"
 Case 1092
  GetLanguageNameLong = "Tatar"
 Case 1098
  GetLanguageNameLong = "Telugu"
 Case 1054
  GetLanguageNameLong = "Thai"
 Case 1055
  GetLanguageNameLong = "Turkish"
 Case 1058
  GetLanguageNameLong = "Ukrainian"
 Case 1056
  GetLanguageNameLong = "Urdu"
 Case 2115
  GetLanguageNameLong = "Uzbek (Cyrillic)"
 Case 1091
  GetLanguageNameLong = "Uzbek (Latin)"
 Case 1066
  GetLanguageNameLong = "Vietnamese"
 Case Else
  GetLanguageNameLong = "Unknown"
 End Select
End Function

'Use GetLanguageNameLong and cut all after first Space
Public Function GetLanguageName(LangId As Long) As String
 GetLanguageName = GetLanguageNameLong(LangId)
On Error Resume Next
If InStr(1, GetLanguageName, " ") <> 0 Then
 GetLanguageName = Left$(GetLanguageName, InStr(1, GetLanguageName, " "))
End If
On Error GoTo 0
End Function


