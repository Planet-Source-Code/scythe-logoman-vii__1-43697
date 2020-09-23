VERSION 5.00
Begin VB.Form FrmCmdlg 
   BorderStyle     =   0  'Kein
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Enabled         =   0   'False
   Icon            =   "FrmCmdlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PicPreview 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   735
      Left            =   240
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmCmdlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Set the new position for preview window
'if the user moves the common dialog we must move too
'This form need to be disabled

Private Sub Timer1_Timer()
 SetWindow
End Sub
