VERSION 5.00
Begin VB.UserControl LMButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ToolboxBitmap   =   "LMButton.ctx":0000
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   1
      X2              =   1
      Y1              =   1
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   1
      X2              =   16
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      X1              =   16
      X2              =   16
      Y1              =   1
      Y2              =   16
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   1
      X2              =   16
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Shape ShpBorder2 
      BorderColor     =   &H80000005&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Shape ShpBorder 
      BorderColor     =   &H00000000&
      Height          =   720
      Left            =   0
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "LMButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'LogoMan Button
'I needed a button that shows a pressed state
'so i wrote one
'its quick coded and not optimized



'Faster than Print
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Get Systemcolors for the Buttonborders.....
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

'This button has only a click event
'i didnt need more
Public Event Click(ButtonDown As Boolean)

'Is this button pressed ?
Dim ButtonDown As Boolean

Private Type POINTAPI
 X As Long
 Y As Long
End Type

Dim ButtonWidth As Long
Dim ButtonHeight As Long
Dim BCL As Long
Dim BCD As Long
Dim BCB As Long
Dim BCC As Long
Dim TL As Long
Dim TT As Long
Dim BtnTxt As String

'Down or not thats the question
'Cahnge state on click
Private Sub UserControl_Click()
 If ButtonDown Then
  ButtonDown = False
 Else
  ButtonDown = True
 End If
 DrawButton
 RaiseEvent Click(ButtonDown)
End Sub

'Startup
'so get the systemcolors
Private Sub UserControl_Initialize()
 BCL = GetSysColor(COLOR_BTNSHADOW)
 BCD = GetSysColor(COLOR_BTNLIGHT)
 BCB = GetSysColor(COLOR_BTNDKSHADOW)
 BCC = GetSysColor(COLOR_BTNHIGHLIGHT)
 BtnTxt = "F"
End Sub

'Resize
'this is the main routine
'it needs some mor api´s for speed
'but its fast enough for me (at the moment)
Private Sub UserControl_Resize()
 UserControl.Cls
 'get the size
 ButtonWidth = UserControl.ScaleWidth - 1
 ButtonHeight = UserControl.ScaleHeight - 1
 'set the new linepositions
 Line1.x2 = ButtonWidth - 1
 Line2.y2 = ButtonHeight - 1
 Line3.x1 = ButtonWidth - 1
 Line3.x2 = ButtonWidth - 1
 Line3.y2 = ButtonHeight - 1
 Line4.x2 = ButtonWidth
 Line4.y1 = ButtonHeight - 1
 Line4.y2 = ButtonHeight - 1
 'Show the shapes
 ShpBorder.Width = ButtonWidth + 1
 ShpBorder.Height = ButtonHeight + 1
 ShpBorder2.Width = ButtonWidth
 ShpBorder2.Height = ButtonHeight
 TT = (ButtonHeight - UserControl.TextHeight(BtnTxt)) / 2
 TL = (ButtonWidth - UserControl.TextWidth(BtnTxt)) / 2
 DrawButton
End Sub

'We have to change the buttons look by state
'so swap colors
Private Sub DrawButton()
 Dim ColorLeft1 As Long
 Dim ColorRight1 As Long
 Dim ColorLeft2 As Long
 Dim ColorRight2 As Long

 If ButtonDown Then
  ColorLeft1 = BCL
  ColorRight1 = BCD
  ColorLeft2 = BCC
  ColorRight2 = BCB
 Else
  BtnVisible = True
  ColorLeft1 = BCD
  ColorRight1 = BCL
  ColorLeft2 = BCB
  ColorRight2 = BCC
 End If

 ShpBorder2.BorderColor = ColorRight2
 ShpBorder.BorderColor = ColorLeft2
 Line1.BorderColor = ColorLeft1
 Line2.BorderColor = ColorLeft1
 Line3.BorderColor = ColorRight1
 Line4.BorderColor = ColorRight1
 TextOut UserControl.hdc, TL, TT, BtnTxt, Len(BtnTxt)

End Sub

'Save the data we set in design time
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
 Call .WriteProperty("Caption", BtnTxt)
 Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
 Call .WriteProperty("Picture", UserControl.Picture)
 End With
End Sub

'Read the data we set in design time
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
 Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
 Set UserControl.Picture = .ReadProperty("Picture", UserControl.Picture)
 BtnTxt = .ReadProperty("Caption", "F")
 End With
End Sub

'The user changed sometihng so update data
Public Property Get Font() As Font
Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal NewFont As Font)
Set UserControl.Font = NewFont
PropertyChanged "Font"
UserControl_Resize
End Property

Public Property Get Caption() As String
Caption = BtnTxt
End Property
Public Property Let Caption(ByVal NewText As String)
BtnTxt = NewText
PropertyChanged "Caption"
UserControl_Resize
End Property

Public Property Get Picture() As Picture
Set Picture = UserControl.Picture
End Property
Public Property Set Picture(ByVal NewPicture As Picture)
Set UserControl.Picture = NewPicture
PropertyChanged "Picture"
UserControl_Resize
End Property

Public Function IsButtonDown() As Boolean
 IsButtonDown = ButtonDown
End Function
Public Function SetButtonstate(Down As Boolean)
 ButtonDown = Down
 UserControl_Resize
End Function
