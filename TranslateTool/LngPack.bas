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

Public Function GetIdFromName(LngName As String) As Long
 Select Case LngName
 Case "Afrikaans"
  GetIdFromName = 1078
 Case "Albanian"
  GetIdFromName = 1052
 Case "Arabic (Algeria)"
  GetIdFromName = 5121
 Case "Arabic (Bahrain)"
  GetIdFromName = 15361
 Case "Arabic (Egypt)"
  GetIdFromName = 3073
 Case "Arabic (Iraq)"
  GetIdFromName = 2049
 Case "Arabic (Jordan)"
  GetIdFromName = 11265
 Case "Arabic (Kuwait)"
  GetIdFromName = 13313
 Case "Arabic (Lebanon)"
  GetIdFromName = 12289
 Case "Arabic (Libya)"
  GetIdFromName = 4097
 Case "Arabic (Morocco)"
  GetIdFromName = 6145
 Case "Arabic (Oman)"
  GetIdFromName = 8193
 Case "Arabic (Qatar)"
  GetIdFromName = 16385
 Case "Arabic (Saudi Arabia)"
  GetIdFromName = 1025
 Case "Arabic (Syria)"
  GetIdFromName = 10241
 Case "Arabic (Tunisia)"
  GetIdFromName = 7169
 Case "Arabic (U.A.E.)"
  GetIdFromName = 14337
 Case "Arabic (Yemen)"
  GetIdFromName = 9217
 Case "Armenian"
  GetIdFromName = 1067
 Case "Azeri (Cyrillic)"
  GetIdFromName = 2092
 Case "Azeri (Latin)"
  GetIdFromName = 1068
 Case "Basque"
  GetIdFromName = 1069
 Case "Belarusian"
  GetIdFromName = 1059
 Case "Bulgarian"
  GetIdFromName = 1026
 Case "Catalan"
  GetIdFromName = 1027
 Case "Chinese (Hong Kong S.A.R.)"
  GetIdFromName = 3076
 Case "Chinese (Macau S.A.R.)"
  GetIdFromName = 5124
 Case "Chinese (PRC)"
  GetIdFromName = 2052
 Case "Chinese (Singapore)"
  GetIdFromName = 4100
 Case "Chinese (Taiwan)"
  GetIdFromName = 1028
 Case "Croatian"
  GetIdFromName = 1050
 Case "Czech"
  GetIdFromName = 1029
 Case "Danish"
  GetIdFromName = 1030
 Case "Divehi"
  GetIdFromName = 1125
 Case "Dutch (Belgium)"
  GetIdFromName = 2067
 Case "Dutch (Netherlands)"
  GetIdFromName = 1043
 Case "English (Australia)"
  GetIdFromName = 3081
 Case "English (Belize)"
  GetIdFromName = 10249
 Case "English (Canada)"
  GetIdFromName = 4105
 Case "English (Caribbean)"
  GetIdFromName = 9225
 Case "English (Ireland)"
  GetIdFromName = 6153
 Case "English (Jamaica)"
  GetIdFromName = 8201
 Case "English (New Zealand)"
  GetIdFromName = 5129
 Case "English (Philippines)"
  GetIdFromName = 13321
 Case "English (South Africa)"
  GetIdFromName = 7177
 Case "English (Trinidad)"
  GetIdFromName = 11273
 Case "English (United Kingdom)"
  GetIdFromName = 2057
 Case "English (United States)"
  GetIdFromName = 1033
 Case "English (Zimbabwe)"
  GetIdFromName = 12297
 Case "Estonian"
  GetIdFromName = 1061
 Case "Faroese"
  GetIdFromName = 1080
 Case "Farsi"
  GetIdFromName = 1065
 Case "Finnish"
  GetIdFromName = 1035
 Case "French (Belgium)"
  GetIdFromName = 2060
 Case "French (Canada)"
  GetIdFromName = 3084
 Case "French (France)"
  GetIdFromName = 1036
 Case "French (Luxembourg)"
  GetIdFromName = 5132
 Case "French (Monaco)"
  GetIdFromName = 6156
 Case "French (Switzerland)"
  GetIdFromName = 4108
 Case "FYRO Macedonian"
  GetIdFromName = 1071
 Case "Galician"
  GetIdFromName = 1110
 Case "Georgian"
  GetIdFromName = 1079
 Case "German (Austria)"
  GetIdFromName = 3079
 Case "German (Germany)"
  GetIdFromName = 1031
 Case "German (Liechtenstein)"
  GetIdFromName = 5127
 Case "German (Luxembourg)"
  GetIdFromName = 4103
 Case "German (Switzerland)"
  GetIdFromName = 2055
 Case "Greek"
  GetIdFromName = 1032
 Case "Gujarati"
  GetIdFromName = 1095
 Case "Hebrew"
  GetIdFromName = 1037
 Case "Hindi"
  GetIdFromName = 1081
 Case "Hungarian"
  GetIdFromName = 1038
 Case "Icelandic"
  GetIdFromName = 1039
 Case "Indonesian"
  GetIdFromName = 1057
 Case "Italian (Italy)"
  GetIdFromName = 1040
 Case "Italian (Switzerland)"
  GetIdFromName = 2064
 Case "Japanese"
  GetIdFromName = 1041
 Case "Kannada"
  GetIdFromName = 1099
 Case "Kazakh"
  GetIdFromName = 1087
 Case "Konkani"
  GetIdFromName = 1111
 Case "Korean"
  GetIdFromName = 1042
 Case "Kyrgyz (Cyrillic)"
  GetIdFromName = 1088
 Case "Latvian"
  GetIdFromName = 1062
 Case "Lithuanian"
  GetIdFromName = 1063
 Case "Malay (Brunei Darussalam)"
  GetIdFromName = 2110
 Case "Malay (Malaysia)"
  GetIdFromName = 1086
 Case "Marathi"
  GetIdFromName = 1102
 Case "Mongolian (Cyrillic)"
  GetIdFromName = 1104
 Case "Norwegian (Bokmal)"
  GetIdFromName = 1044
 Case "Norwegian (Nynorsk)"
  GetIdFromName = 2068
 Case "Polish"
  GetIdFromName = 1045
 Case "Portuguese (Brazil)"
  GetIdFromName = 1046
 Case "Portuguese (Portugal)"
  GetIdFromName = 2070
 Case "Punjabi"
  GetIdFromName = 1094
 Case "Romanian"
  GetIdFromName = 1048
 Case "Russian"
  GetIdFromName = 1049
 Case "Sanskrit"
  GetIdFromName = 1103
 Case "Serbian (Cyrillic)"
  GetIdFromName = 3098
 Case "Serbian (Latin)"
  GetIdFromName = 2074
 Case "Slovak"
  GetIdFromName = 1051
 Case "Slovenian"
  GetIdFromName = 1060
 Case "Spanish (Argentina)"
  GetIdFromName = 11274
 Case "Spanish (Bolivia)"
  GetIdFromName = 16394
 Case "Spanish (Chile)"
  GetIdFromName = 13322
 Case "Spanish (Colombia)"
  GetIdFromName = 9226
 Case "Spanish (Costa Rica)"
  GetIdFromName = 5130
 Case "Spanish (Dominican Republic)"
  GetIdFromName = 7178
 Case "Spanish (Ecuador)"
  GetIdFromName = 12298
 Case "Spanish (El Salvador)"
  GetIdFromName = 17418
 Case "Spanish (Guatemala)"
  GetIdFromName = 4106
 Case "Spanish (Honduras)"
  GetIdFromName = 18442
 Case "Spanish (International Sort)"
  GetIdFromName = 3082
 Case "Spanish (Mexico)"
  GetIdFromName = 2058
 Case "Spanish (Nicaragua)"
  GetIdFromName = 19466
 Case "Spanish (Panama)"
  GetIdFromName = 6154
 Case "Spanish (Paraguay)"
  GetIdFromName = 15370
 Case "Spanish (Peru)"
  GetIdFromName = 10250
 Case "Spanish (Puerto Rico)"
  GetIdFromName = 20490
 Case "Spanish (Traditional Sort)"
  GetIdFromName = 1034
 Case "Spanish (Uruguay)"
  GetIdFromName = 14346
 Case "Spanish (Venezuela)"
  GetIdFromName = 8202
 Case "Swahili"
  GetIdFromName = 1089
 Case "Swedish"
  GetIdFromName = 1053
 Case "Swedish (Finland)"
  GetIdFromName = 2077
 Case "Syriac"
  GetIdFromName = 1114
 Case "Tamil"
  GetIdFromName = 1097
 Case "Tatar"
  GetIdFromName = 1092
 Case "Telugu"
  GetIdFromName = 1098
 Case "Thai"
  GetIdFromName = 1054
 Case "Turkish"
  GetIdFromName = 1055
 Case "Ukrainian"
  GetIdFromName = 1058
 Case "Urdu"
  GetIdFromName = 1056
 Case "Uzbek (Cyrillic)"
  GetIdFromName = 2115
 Case "Uzbek (Latin)"
  GetIdFromName = 1091
 Case "Vietnamese"
  GetIdFromName = 1066
 Case "Unknown"
  GetIdFromName = 0
 End Select
End Function

