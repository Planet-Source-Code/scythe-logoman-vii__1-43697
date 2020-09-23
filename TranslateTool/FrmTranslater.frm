VERSION 5.00
Begin VB.Form FrmTranslator 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Translator"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   Icon            =   "FrmTranslater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame FrmNorm 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox CboControl 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.Frame FrmCombo 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame1"
         Height          =   3255
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   5535
         Begin VB.ComboBox CboCbo 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox TxtCbo 
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Type your translation in here"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label5 
            Caption         =   "Listentry"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CheckBox ChckAuto 
         Caption         =   "Autotranslate double"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox TxtToolTipN 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   3720
         Width           =   5415
      End
      Begin VB.TextBox TxtTitelN 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   2040
         Width           =   5415
      End
      Begin VB.TextBox TxtToolTipO 
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3000
         Width           =   5415
      End
      Begin VB.TextBox TxtTitelO 
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "Type your translation in here"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Type your translation in here"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Control"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "ToolTip"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Titel"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.ComboBox CboNew 
      Height          =   315
      ItemData        =   "FrmTranslater.frx":030A
      Left            =   2880
      List            =   "FrmTranslater.frx":04A3
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox CboOriginal 
      Height          =   315
      ItemData        =   "FrmTranslater.frx":0D05
      Left            =   120
      List            =   "FrmTranslater.frx":0D07
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Zentriert
      Caption         =   "Please send new Language Packs to scythe@cablenet.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   6600
      Width           =   5775
   End
   Begin VB.Label LblLngID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "New"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Original"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "FrmTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Translator by Scythe@cablenet.de
'Use this little tool to translate Language Packs
'Think thats easyer than doing it from hand using edit

'No comments because this is to easy to understand


'To use it in your own Apps
'incl. Module LngPack.bas form PSC or better newer version form LogoMan
'Change this line in Form_Load
'LngDir = LngDir & "LangPacks\"
'            to
'LngDir = LngDir & ??MyLanguageDirectoy??
            

Option Explicit


Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
Private Const LOCALE_ILANGUAGE = &H1

Private Type Lng
 ToolTip As String
 Caption As String
 List() As String
 CtrlType As String
 Idx As String
 CtrlName As String
 ShownName As String
 Counter As Long
 Header As String
End Type



Dim OrgPack(200) As Lng
Dim NewPack(200) As Lng
Dim LngDir As String

Private Sub CboCbo_Click()
 Dim X As Long
 X = CboCbo.ListIndex
 If X = -1 Then Exit Sub
 ReDim Preserve NewPack(CboControl.ListIndex).List(UBound(OrgPack(CboControl.ListIndex).List()))
 If X = CboCbo.ListCount - 1 Then
  TxtCbo.Enabled = False
  TxtCbo.Text = ""
 Else
  TxtCbo.Enabled = True
  TxtCbo.Text = NewPack(CboControl.ListIndex).List(X)
 End If
End Sub




Private Sub Form_Terminate()
 Unload Me
 End
End Sub

Private Sub TxtCbo_Change()
 NewPack(CboControl.ListIndex).List(CboCbo.ListIndex) = TxtCbo.Text
 Autotranslate OrgPack(CboControl.ListIndex).List(CboCbo.ListIndex), TxtCbo.Text
End Sub

Private Sub CboControl_Click()
 Dim X As Long
 Dim i As Long
 X = CboControl.ListIndex
 If X = -1 Then Exit Sub


 If OrgPack(X).CtrlType <> "#" Then 'This is no Combobox
  FrmCombo.Visible = False
  If OrgPack(X).Caption <> "" Then
   TxtTitelO.Text = OrgPack(X).Caption
   TxtTitelN.Text = NewPack(X).Caption
   TxtTitelN.Enabled = True
  Else
   TxtTitelN.Enabled = False
   TxtTitelN.Text = ""
  End If
  If OrgPack(X).ToolTip <> "" Then
   TxtToolTipO.Text = OrgPack(X).ToolTip
   TxtToolTipN.Text = NewPack(X).ToolTip
   TxtToolTipN.Enabled = True
  Else
   TxtToolTipN.Enabled = False
   TxtToolTipO.Text = ""
  End If
 Else
  For i = 0 To UBound(OrgPack(X).List())
   CboCbo.AddItem (OrgPack(X).List(i))
  Next i
  CboCbo.ListIndex = 0
  FrmCombo.Visible = True
 End If
End Sub
Private Sub TxtTitelN_Change()
 NewPack(CboControl.ListIndex).Caption = TxtTitelN.Text
 Autotranslate OrgPack(CboControl.ListIndex).Caption, TxtTitelN.Text
End Sub
Private Sub TxtToolTipN_Change()
 NewPack(CboControl.ListIndex).ToolTip = TxtToolTipN.Text
 Autotranslate OrgPack(CboControl.ListIndex).ToolTip, TxtToolTipN.Text
End Sub


Private Sub CboNew_Click()
 Dim Fname As String
 Dim i As Long
 If CboNew.ListIndex = -1 Then Exit Sub
 Fname = Trim(Str(GetIdFromName(CboNew.List(CboNew.ListIndex))) & ".lng")
 If LenB(Dir$(LngDir & Fname)) <> 0 Then
  LoadLanguage LngDir & Fname, NewPack()
  Do Until NewPack(i).ShownName = ""
   CboControl.AddItem NewPack(i).ShownName
   i = i + 1
  Loop
  CboControl.ListIndex = 0
  CboControl_Click
 End If
 FrmNorm.Visible = True
End Sub

Private Sub CboOriginal_Click()
 Dim Fname As String
 Dim i As Long
 If CboOriginal.ListIndex = -1 Then Exit Sub
 Fname = Trim(Str(GetIdFromName(CboOriginal.List(CboOriginal.ListIndex))) & ".lng")

 LoadLanguage LngDir & Fname, OrgPack()
 Do Until OrgPack(i).ShownName = ""
  CboControl.AddItem OrgPack(i).ShownName
  i = i + 1
 Loop
 CboControl.ListIndex = 0
 CboControl_Click
End Sub

Private Sub Form_Load()
 Dim Fname As String
 Dim h As Long

 'Get the LanguagePack directory
 LngDir = App.Path
 If Right$(LngDir, 1) <> "\" Then LngDir = LngDir & "\"
 '********************************
 '*  Change this for ur Own Apps *
 '********************************
 LngDir = LngDir & "LangPacks\"



 'Get all installed Languages
 Fname = Dir$(LngDir & "*.lng")

 If Fname = "" Then
  MsgBox "Cant find Language Packs or Directory" & vbCrLf & LngDir, vbCritical, "Search for existing Language Packs"
  Unload Me
  End
 End If

 Do Until Fname = ""
  h = Val(Left$(Fname, Len(Fname) - 3))
  Fname = Trim(GetLanguageNameLong(h))
  CboOriginal.AddItem Fname
  Fname = Dir$
 Loop
 CboOriginal.ListIndex = 0
 CboNew.ListIndex = -1

 'Show windowslanguage
 LblLngID.Caption = "Your Windows Homelanguage is " & GetLanguageNameLong(GetLanguageID)
End Sub

'Load LanguagePack into Array
Private Function LoadLanguage(LanguageFile As String, ByRef Pack() As Lng)
 Dim i As Long
 Dim Ctr As Long
 Dim Tmp As String
 Dim CtrlName As String
 Dim LngCtr As Long

 'Open the LanguageFile
 Open LanguageFile For Input As #1
  Do Until EOF(1)

   Line Input #1, Tmp       'Get the Header
   Pack(LngCtr).Header = Tmp

   If Len(Tmp) > 1 Then     'Extract the counter if there is one
    Pack(LngCtr).Counter = Val(Right$(Tmp, Len(Tmp) - 1))
   End If

   Line Input #1, Pack(LngCtr).CtrlName 'Get ControlName

   Line Input #1, Pack(LngCtr).Idx      'Get Index

   'Show the real name of the Control
   If Pack(LngCtr).Idx = "-1" Then
    Pack(LngCtr).ShownName = Pack(LngCtr).CtrlName
   Else
    Pack(LngCtr).ShownName = Pack(LngCtr).CtrlName & "(" & Pack(LngCtr).Idx & ")"
   End If

   Pack(LngCtr).CtrlType = Left$(Tmp, 1)

   Select Case Left$(Tmp, 1)


   Case "#" 'Combobox
    ReDim Pack(LngCtr).List(Pack(LngCtr).Counter)
    For i = 0 To Pack(LngCtr).Counter 'Now read all lines for the Combobox
     Line Input #1, Pack(LngCtr).List(i)
    Next i

   Case "~" 'Tabstrip
    'we create a new control for every Tab
    For i = 0 To Pack(LngCtr).Counter - 1
     Line Input #1, Pack(LngCtr).Caption
     Line Input #1, Pack(LngCtr).ToolTip
     LngCtr = LngCtr + 1
     Pack(LngCtr).Idx = Pack(LngCtr - 1).Idx
     Pack(LngCtr).ShownName = Pack(LngCtr - 1).ShownName
    Next i
    LngCtr = LngCtr - 1

   Case "+" 'ListView
    'Same as Tabstrip
    For i = 0 To Pack(LngCtr).Counter - 1
     Line Input #1, Pack(LngCtr).Caption
     LngCtr = LngCtr + 1
     Pack(LngCtr).Idx = Pack(LngCtr - 1).Idx
     Pack(LngCtr).ShownName = Pack(LngCtr - 1).ShownName
    Next i
    LngCtr = LngCtr - 1

   Case "^" 'Control with Only Caption
    Line Input #1, Pack(LngCtr).Caption

   Case "°" 'Control with Only Tooltip
    Line Input #1, Pack(LngCtr).ToolTip

   Case "*" 'Control with Caption & ToolTipText
    Line Input #1, Pack(LngCtr).Caption
    Line Input #1, Pack(LngCtr).ToolTip

   End Select
   LngCtr = LngCtr + 1
  Loop
 Close
End Function

'Save new LanguagePack
Private Sub CmdSave_Click()
 Dim Fname As String
 Dim i As Long
 Dim f As Long
 Dim NoTranslate As Boolean
 Dim LngCtr As Long

 If CboNew.ListIndex = -1 Then Exit Sub
 Fname = Trim(Str(GetIdFromName(CboNew.List(CboNew.ListIndex))) & ".lng")

 Do Until NewPack(i).ShownName = ""
  If OrgPack(i).CtrlType <> "#" Then
   If OrgPack(i).Caption <> "" Then
    If NewPack(i).Caption = "" Then NoTranslate = True
   End If
   If OrgPack(i).ToolTip <> "" Then
    If NewPack(i).ToolTip = "" Then NoTranslate = True
   End If

  Else
   ReDim Preserve NewPack(i).List(UBound(OrgPack(i).List()))
   For f = 0 To UBound(OrgPack(i).List()) - 1
    If NewPack(i).List(f) = "" Then NoTranslate = True
   Next f
  End If
  i = i + 1
 Loop

 If NoTranslate Then
  If MsgBox("There are still some translations missing." & vbCrLf & "Are you sure you want to save ?", vbCritical + vbOKCancel + vbDefaultButton2, "Save New Languagepack") = vbCancel Then Exit Sub
 End If
 If LenB(Dir$(LngDir & Fname)) <> 0 Then
  If MsgBox("The language pack still exists." & vbCrLf & "Are you sure to overwrite language pack ?", vbCritical + vbOKCancel + vbDefaultButton2, "Save New Languagepack") = vbCancel Then Exit Sub
 End If



 Open LngDir & Fname For Output As #1
  Do Until NewPack(LngCtr).ShownName = ""
   Print #1, NewPack(LngCtr).Header
   Print #1, NewPack(LngCtr).CtrlName
   Print #1, NewPack(LngCtr).Idx

   Select Case NewPack(LngCtr).CtrlType

   Case "#" 'Combobox
    For i = 0 To NewPack(LngCtr).Counter
     Print #1, NewPack(LngCtr).List(i)
    Next i

   Case "~" 'Tabstrip
    For i = 0 To NewPack(LngCtr).Counter - 1
     Print #1, NewPack(LngCtr).Caption
     Print #1, NewPack(LngCtr).ToolTip
     LngCtr = LngCtr + 1
    Next i
    LngCtr = LngCtr - 1

   Case "+" 'ListView
    For i = 0 To NewPack(LngCtr).Counter - 1
     Print #1, NewPack(LngCtr).Caption
     LngCtr = LngCtr + 1
    Next i
    LngCtr = LngCtr - 1

   Case "^" 'Control with Only Caption
    Print #1, NewPack(LngCtr).Caption

   Case "°" 'Control with Only Tooltip
    Print #1, NewPack(LngCtr).ToolTip

   Case "*" 'Control with Caption & ToolTipText
    Print #1, NewPack(LngCtr).Caption
    Print #1, NewPack(LngCtr).ToolTip

   End Select
   LngCtr = LngCtr + 1
  Loop
 Close
 MsgBox "Saved Language Pack as" & vbCrLf & Fname, , "Save new Translation"
End Sub

'Automatic Translate every Cption,Tooltext...
'if the Original is found a second time
Public Sub Autotranslate(OrgText As String, NewText As String)
 If ChckAuto.Value = 0 Then Exit Sub
 Dim i As Long
 Dim f As Long
 If NewText = "" Then Exit Sub
 Do Until NewPack(i).ShownName = ""
  If OrgPack(i).CtrlType <> "#" Then
   If OrgPack(i).Caption = OrgText Then NewPack(i).Caption = NewText
   If OrgPack(i).ToolTip = OrgText Then NewPack(i).ToolTip = NewText
  Else
   ReDim Preserve NewPack(i).List(UBound(OrgPack(i).List()))
   For f = 0 To UBound(OrgPack(i).List()) - 1
    If OrgPack(i).List(f) = OrgText Then NewPack(i).List(f) = NewText
   Next f
  End If
  i = i + 1
 Loop
End Sub


Private Function GetLanguageID() As Long
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
Private Function GetLanguageNameLong(LangId As Long) As String
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

'New function to get the Id form a Language Name
Private Function GetIdFromName(LngName As String) As Long
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
 End Select
End Function
