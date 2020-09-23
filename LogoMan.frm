VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form LogoMan 
   Caption         =   "LogoMan V7 by Scythe   scythe@cablenet.de"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "LogoMan.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "LogoMan.frx":11C2
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrmMainTab 
      BorderStyle     =   0  'Kein
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   8760
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2775
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   720
         Width           =   8535
         Begin VB.CheckBox ChkTex 
            Caption         =   "Texture"
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   133
            ToolTipText     =   "Use deforming Textures"
            Top             =   120
            Width           =   1455
         End
         Begin VB.HScrollBar HScrDL 
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   6960
            Max             =   100
            Min             =   -100
            SmallChange     =   5
            TabIndex        =   35
            Top             =   1920
            Value           =   100
            Width           =   1455
         End
         Begin VB.CommandButton CmdToTex 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   5520
            TabIndex        =   34
            ToolTipText     =   "Copy loaded Texture to Texturepath"
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   33
            ToolTipText     =   "Load Texture from any Drive"
            Top             =   1920
            Width           =   1275
         End
         Begin VB.FileListBox File 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   2
            Left            =   4200
            Pattern         =   "*.jpg;*.gif;*.bmp"
            TabIndex        =   32
            ToolTipText     =   "Texturenames"
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton CmdDLOK 
            Caption         =   "OK"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   31
            Top             =   1920
            Width           =   375
         End
         Begin VB.HScrollBar HScrDL 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   2640
            Max             =   100
            Min             =   -100
            SmallChange     =   5
            TabIndex        =   30
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton CmdToTex 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   29
            ToolTipText     =   "Copy loaded Texture to Texturepath"
            Top             =   1920
            Width           =   1155
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Load Texture from any Drive"
            Top             =   1920
            Width           =   1155
         End
         Begin VB.FileListBox File 
            Height          =   1455
            Index           =   0
            Left            =   120
            Pattern         =   "*.jpg;*.gif;*.bmp"
            TabIndex        =   27
            ToolTipText     =   "Texturenames"
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Picture"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   1455
         End
         Begin VB.Image ImgPrev 
            Height          =   1455
            Index           =   0
            Left            =   2640
            Stretch         =   -1  'True
            ToolTipText     =   "Preview"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image ImgPrev 
            Height          =   1455
            Index           =   2
            Left            =   6960
            Stretch         =   -1  'True
            ToolTipText     =   "Preview"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2415
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CommandButton CmdSolCol 
            Appearance      =   0  '2D
            BackColor       =   &H000000FF&
            Height          =   375
            Index           =   0
            Left            =   1080
            MaskColor       =   &H00000000&
            Style           =   1  'Grafisch
            TabIndex        =   40
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2415
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   3600
            Style           =   1  'Grafisch
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
         Begin VB.ComboBox CboGrStyle 
            Height          =   315
            Index           =   0
            ItemData        =   "LogoMan.frx":1504
            Left            =   2640
            List            =   "LogoMan.frx":1511
            Style           =   2  'Dropdown-Liste
            TabIndex        =   44
            ToolTipText     =   "Gradientstyle"
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton CmdSwapColors 
            Caption         =   "<->"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   4.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   43
            ToolTipText     =   "Swap Colors"
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H00CC0000&
            Height          =   375
            Index           =   0
            Left            =   2640
            Style           =   1  'Grafisch
            TabIndex        =   42
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Transparent"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   38
         ToolTipText     =   "Invisible font and normal border"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Gradient"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   37
         ToolTipText     =   "Multicolor Fonts"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Solid Color"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   25
         ToolTipText     =   "Single Color Font"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Texture"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   24
         ToolTipText     =   "Use picture Overlays"
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   17
         Text            =   "LogoMan"
         ToolTipText     =   "Type Text in here"
         Top             =   0
         Width           =   2775
      End
      Begin VB.ComboBox CboName 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   16
         ToolTipText     =   "Fontname"
         Top             =   15
         Width           =   2100
      End
      Begin VB.ComboBox CboSize 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "Fontsize"
         Top             =   15
         Width           =   675
      End
      Begin PrjLogoMan.LMButton LMBAlign 
         Height          =   300
         Index           =   0
         Left            =   4080
         TabIndex        =   14
         ToolTipText     =   "Align Left"
         Top             =   45
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":1535
      End
      Begin PrjLogoMan.LMButton LMBStytle 
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   18
         ToolTipText     =   "Bold"
         Top             =   45
         Width           =   300
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   "B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":1B17
      End
      Begin PrjLogoMan.LMButton LMBStytle 
         Height          =   300
         Index           =   1
         Left            =   3240
         TabIndex        =   19
         ToolTipText     =   "Kursiv"
         Top             =   45
         Width           =   300
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   "K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":1B33
      End
      Begin PrjLogoMan.LMButton LMBStytle 
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   20
         ToolTipText     =   "Underline"
         Top             =   45
         Width           =   300
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   "U"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":1B4F
      End
      Begin PrjLogoMan.LMButton LMBAlign 
         Height          =   300
         Index           =   1
         Left            =   4440
         TabIndex        =   21
         ToolTipText     =   "Align Zenter"
         Top             =   45
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":1B6B
      End
      Begin PrjLogoMan.LMButton LMBAlign 
         Height          =   300
         Index           =   2
         Left            =   4800
         TabIndex        =   22
         ToolTipText     =   "Align Right"
         Top             =   45
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":214D
      End
      Begin PrjLogoMan.LMButton LMBAlign 
         Height          =   300
         Index           =   3
         Left            =   5160
         TabIndex        =   23
         ToolTipText     =   "Choose ur textposition"
         Top             =   45
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "LogoMan.frx":272F
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2535
         Index           =   3
         Left            =   0
         TabIndex        =   118
         Top             =   960
         Visible         =   0   'False
         Width           =   8535
      End
   End
   Begin VB.Frame FrmMainTab 
      BorderStyle     =   0  'Kein
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   46
      Top             =   360
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2655
         Index           =   5
         Left            =   0
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   8895
         Begin VB.CommandButton CmdSolCol 
            Appearance      =   0  '2D
            BackColor       =   &H000000FF&
            Height          =   315
            Index           =   1
            Left            =   1680
            MaskColor       =   &H00000000&
            Style           =   1  'Grafisch
            TabIndex        =   68
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox ChkGlow 
            Caption         =   "Glow"
            Height          =   255
            Left            =   1680
            TabIndex        =   67
            ToolTipText     =   "Create glowing Borders"
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtBordersize 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "2"
         ToolTipText     =   "Size of the Border"
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox ChkBorder 
         Caption         =   "Border"
         Height          =   255
         Left            =   0
         TabIndex        =   121
         ToolTipText     =   "Border arround the Font"
         Top             =   120
         Width           =   1005
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Max             =   -1
         Min             =   -8
         TabIndex        =   120
         Top             =   120
         Value           =   -2
         Width           =   135
      End
      Begin VB.CheckBox ChkRight 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6960
         TabIndex        =   63
         Top             =   420
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox ChkBottom 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6780
         TabIndex        =   65
         Top             =   600
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox ChkTop 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6780
         TabIndex        =   64
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox ChkLeft 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6600
         TabIndex        =   66
         Top             =   420
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Texture"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   61
         ToolTipText     =   "Use picture Overlays"
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Solid Color"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   60
         ToolTipText     =   "Single Color Font"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Gradient"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   48
         ToolTipText     =   "Multicolor Fonts"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Fade"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   47
         ToolTipText     =   "Fade between Font and Background"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2655
         Index           =   7
         Left            =   0
         TabIndex        =   74
         Top             =   720
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2775
         Index           =   4
         Left            =   0
         TabIndex        =   49
         Top             =   720
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CheckBox ChkTex 
            Caption         =   "Texture"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   132
            ToolTipText     =   "Use deforming Textures"
            Top             =   120
            Width           =   1455
         End
         Begin VB.FileListBox File 
            Height          =   1455
            Index           =   1
            Left            =   120
            Pattern         =   "*.jpg;*.gif;*.bmp"
            TabIndex        =   58
            ToolTipText     =   "Texturenames"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Load Texture from any Drive"
            Top             =   1920
            Width           =   1155
         End
         Begin VB.CommandButton CmdToTex 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   56
            ToolTipText     =   "Copy loaded Texture to Texturepath"
            Top             =   1920
            Width           =   1155
         End
         Begin VB.HScrollBar HScrDL 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   2640
            Max             =   100
            Min             =   -100
            SmallChange     =   5
            TabIndex        =   55
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton CmdDLOK 
            Caption         =   "OK"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   54
            Top             =   1920
            Width           =   375
         End
         Begin VB.FileListBox File 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   3
            Left            =   4200
            Pattern         =   "*.jpg;*.gif;*.bmp"
            TabIndex        =   53
            ToolTipText     =   "Texturenames"
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   52
            ToolTipText     =   "Load Texture from any Drive"
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton CmdToTex 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   5520
            TabIndex        =   51
            ToolTipText     =   "Copy loaded Texture to Texturepath"
            Top             =   1920
            Width           =   1275
         End
         Begin VB.HScrollBar HScrDL 
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   6960
            Max             =   100
            Min             =   -100
            SmallChange     =   5
            TabIndex        =   50
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Image ImgPrev 
            Height          =   1455
            Index           =   3
            Left            =   6960
            Stretch         =   -1  'True
            ToolTipText     =   "Preview"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image ImgPrev 
            Height          =   1455
            Index           =   1
            Left            =   2640
            Stretch         =   -1  'True
            ToolTipText     =   "Preview"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Picture"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2655
         Index           =   6
         Left            =   0
         TabIndex        =   69
         Top             =   720
         Visible         =   0   'False
         Width           =   8895
         Begin VB.ComboBox CboGrStyle 
            Height          =   315
            Index           =   1
            ItemData        =   "LogoMan.frx":2D11
            Left            =   3240
            List            =   "LogoMan.frx":2D1E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   73
            ToolTipText     =   "Gradientstyle"
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   3240
            Style           =   1  'Grafisch
            TabIndex        =   72
            Top             =   375
            Width           =   495
         End
         Begin VB.CommandButton CmdSwapColors 
            Caption         =   "<->"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   4.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3720
            TabIndex        =   71
            ToolTipText     =   "Swap Colors"
            Top             =   375
            Width           =   495
         End
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H000000FF&
            Height          =   375
            Index           =   3
            Left            =   4200
            Style           =   1  'Grafisch
            TabIndex        =   70
            Top             =   375
            Width           =   495
         End
      End
   End
   Begin VB.Frame FrmMainTab 
      BorderStyle     =   0  'Kein
      Height          =   3255
      Index           =   3
      Left            =   120
      TabIndex        =   95
      Top             =   360
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox ChkLight 
         Caption         =   "Backlight"
         Height          =   255
         Left            =   6960
         TabIndex        =   119
         ToolTipText     =   "Turn Backgroundlight ON/OFF"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Frame FrmLight 
         Height          =   2895
         Left            =   6840
         TabIndex        =   110
         Top             =   120
         Width           =   1695
         Begin VB.CheckBox ChkReal 
            Caption         =   "Realism"
            Height          =   240
            Left            =   120
            TabIndex        =   114
            ToolTipText     =   "Turn realistic lighteffect on/off"
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton CmdLightCol 
            Appearance      =   0  '2D
            BackColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            MaskColor       =   &H00000000&
            Style           =   1  'Grafisch
            TabIndex        =   113
            Top             =   360
            Width           =   1335
         End
         Begin VB.HScrollBar HScrStrnght 
            Height          =   255
            LargeChange     =   2
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   112
            Top             =   2040
            Value           =   10
            Width           =   1335
         End
         Begin VB.HScrollBar HScrDist 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   199
            Min             =   1
            TabIndex        =   111
            Top             =   1440
            Value           =   1
            Width           =   1335
         End
         Begin VB.Label LblStrenght 
            Caption         =   "Strenght"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label LblDistance 
            Caption         =   "Distance"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.CommandButton CmdHover 
         Caption         =   "Hover"
         Height          =   375
         Left            =   1680
         TabIndex        =   109
         ToolTipText     =   "Windows calls this Shadow"
         Top             =   1860
         Width           =   1815
      End
      Begin VB.CommandButton CmdShadow 
         Caption         =   "Shadow"
         Height          =   375
         Left            =   1680
         TabIndex        =   108
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdRedraw 
         Caption         =   "Undo All"
         Height          =   375
         Left            =   1920
         TabIndex        =   107
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Deform"
         Height          =   1335
         Left            =   1680
         TabIndex        =   104
         Top             =   0
         Width           =   1815
         Begin VB.CommandButton CmdBend 
            Caption         =   "Bend Vertical"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   106
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton CmdBend 
            Caption         =   "Bend Horizental"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Freehand"
         Height          =   2415
         Left            =   0
         TabIndex        =   98
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton CmdPoint 
            Caption         =   "Point"
            Height          =   420
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Width           =   1260
         End
         Begin VB.CommandButton CmdLine 
            Caption         =   "Line"
            Height          =   420
            Left            =   120
            TabIndex        =   102
            ToolTipText     =   "Line"
            Top             =   840
            Width           =   1260
         End
         Begin VB.CommandButton CmdBox 
            Caption         =   "Box"
            Height          =   420
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   101
            ToolTipText     =   "Box"
            Top             =   1320
            Width           =   1260
         End
         Begin VB.CommandButton CmdPip 
            Caption         =   "Pic Color"
            Height          =   420
            Left            =   120
            TabIndex        =   100
            ToolTipText     =   "Pipett"
            Top             =   1800
            Width           =   780
         End
         Begin VB.CommandButton CmdDrawColor 
            BackColor       =   &H00000000&
            Height          =   420
            Left            =   885
            Style           =   1  'Grafisch
            TabIndex        =   99
            Top             =   1800
            Width           =   495
         End
      End
      Begin VB.ComboBox CboFX 
         Height          =   315
         ItemData        =   "LogoMan.frx":2D42
         Left            =   120
         List            =   "LogoMan.frx":2D52
         Style           =   2  'Dropdown-Liste
         TabIndex        =   97
         Top             =   2760
         Width           =   1215
      End
      Begin VB.PictureBox PicMagnify 
         BackColor       =   &H00FFFFFF&
         ClipControls    =   0   'False
         DrawMode        =   7  'Invers
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   3600
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   96
         Top             =   120
         Width           =   3060
         Begin VB.Shape ShpSun 
            BackColor       =   &H0000FFFF&
            BorderColor     =   &H0000FFFF&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Ausgefüllt
            Height          =   255
            Left            =   0
            Shape           =   3  'Kreis
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Filters"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Frame FrmMainTab 
      BorderStyle     =   0  'Kein
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2775
         Index           =   8
         Left            =   120
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CheckBox ChkStretch 
            Caption         =   "Stretch"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   87
            ToolTipText     =   "Stretch Picture to Logosize"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton CmdBackground 
            Caption         =   "Load"
            Height          =   375
            Left            =   0
            TabIndex        =   86
            ToolTipText     =   "Load a Picture"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CheckBox ChkShrink 
            Caption         =   "Shrink"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   85
            ToolTipText     =   "Shrink Picture to Logosize"
            Top             =   870
            Width           =   1335
         End
         Begin VB.CheckBox ChkTile 
            Caption         =   "Tile"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   84
            ToolTipText     =   "Tile Picture over Logo"
            Top             =   1140
            Width           =   1335
         End
         Begin VB.CommandButton CmdClearBack 
            Caption         =   "Clear"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1110
            TabIndex        =   83
            ToolTipText     =   "Disable Picture"
            Top             =   120
            Width           =   660
         End
         Begin VB.HScrollBar HScrMoveBack 
            Height          =   255
            Left            =   30
            TabIndex        =   82
            Top             =   1800
            Width           =   735
         End
         Begin VB.VScrollBar VScrMoveBack 
            Height          =   735
            Left            =   270
            TabIndex        =   81
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton VmdTmp 
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2775
         Index           =   9
         Left            =   0
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CommandButton CmdBackcolor 
            Caption         =   "Color"
            Height          =   375
            Left            =   1560
            TabIndex        =   89
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Gradient"
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   79
         ToolTipText     =   "Multicolor Background"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Solid Color"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   77
         ToolTipText     =   "Single Color Background"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTab0 
         Caption         =   "Picture"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   76
         ToolTipText     =   "Use a Backgoundpicture"
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame FrmSub0 
         BorderStyle     =   0  'Kein
         Height          =   2775
         Index           =   10
         Left            =   120
         TabIndex        =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H00C00000&
            Height          =   375
            Index           =   4
            Left            =   3120
            Style           =   1  'Grafisch
            TabIndex        =   94
            Top             =   495
            Width           =   495
         End
         Begin VB.CommandButton CmdGrCol 
            BackColor       =   &H000000FF&
            Height          =   375
            Index           =   5
            Left            =   4080
            Style           =   1  'Grafisch
            TabIndex        =   93
            Top             =   495
            Width           =   495
         End
         Begin VB.ComboBox CboGrStyle 
            Height          =   315
            Index           =   2
            ItemData        =   "LogoMan.frx":2D7B
            Left            =   3120
            List            =   "LogoMan.frx":2D88
            Style           =   2  'Dropdown-Liste
            TabIndex        =   92
            ToolTipText     =   "Gradientstyle"
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton CmdSwapColors 
            Caption         =   "<->"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   4.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   3600
            TabIndex        =   91
            ToolTipText     =   "Swap Colors"
            Top             =   495
            Width           =   495
         End
      End
   End
   Begin VB.Frame FrmSettings 
      Height          =   1095
      Left            =   3120
      TabIndex        =   123
      Top             =   3750
      Width           =   5880
      Begin VB.Frame FrmEmEnLogo 
         Height          =   1095
         Left            =   3360
         TabIndex        =   145
         Top             =   0
         Width           =   2535
         Begin VB.VScrollBar VScrEmEn 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1700
            Max             =   0
            Min             =   -8
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   540
            Value           =   -2
            Width           =   135
         End
         Begin VB.CheckBox ChkEmEnBox 
            Caption         =   "Box"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1440
            TabIndex        =   149
            ToolTipText     =   "Box or Outline"
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox ChkEmbosFont 
            Caption         =   "Embos"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            ToolTipText     =   "Embos the Font"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox TxtEmEn 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   147
            Text            =   "2"
            ToolTipText     =   "Points away from Font"
            Top             =   540
            Width           =   375
         End
         Begin VB.CheckBox ChkEngraveFont 
            Caption         =   "Engrave"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            ToolTipText     =   "Engrave the Font"
            Top             =   600
            Width           =   1035
         End
      End
      Begin VB.CheckBox ChkInvert 
         Caption         =   "Invert"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox ChkMirror 
         Caption         =   "Mirror"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         ToolTipText     =   "Create Mirror Fonts"
         Top             =   240
         Width           =   1095
      End
      Begin VB.VScrollBar VScrAlpha 
         Height          =   285
         Left            =   3120
         Max             =   -1
         Min             =   -99
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   540
         Value           =   -50
         Width           =   135
      End
      Begin VB.CheckBox ChkAlpha 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   1440
         TabIndex        =   126
         ToolTipText     =   "Transparency"
         Top             =   600
         Width           =   1275
      End
      Begin VB.CheckBox ChkAlias 
         Caption         =   "Anti Alias"
         Height          =   255
         Left            =   1440
         TabIndex        =   125
         ToolTipText     =   "Softter Fonts"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtAlpha 
         Height          =   285
         Left            =   2760
         TabIndex        =   124
         Text            =   "50"
         ToolTipText     =   "Transparency in %"
         Top             =   540
         Width           =   375
      End
   End
   Begin VB.Frame FrmStd 
      Height          =   1095
      Left            =   0
      TabIndex        =   136
      Top             =   3750
      Width           =   3135
      Begin VB.CommandButton CmdClipboard 
         Caption         =   "Copy to Clipboard"
         Height          =   525
         Index           =   0
         Left            =   1560
         TabIndex        =   144
         ToolTipText     =   "Copy Logo to Clipboard"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame AntiScrollMouse 
         BorderStyle     =   0  'Kein
         Height          =   855
         Left            =   120
         TabIndex        =   137
         Top             =   180
         Width           =   1455
         Begin VB.VScrollBar VscrX 
            Height          =   285
            Left            =   1200
            Max             =   -20
            Min             =   -2000
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   120
            Value           =   -20
            Width           =   135
         End
         Begin VB.VScrollBar VscrY 
            Height          =   285
            Left            =   1200
            Max             =   -20
            Min             =   -2000
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   480
            Value           =   -20
            Width           =   135
         End
         Begin VB.TextBox TxtSizeY 
            Height          =   285
            Left            =   720
            TabIndex        =   139
            Text            =   "2"
            ToolTipText     =   "Vertical Logosize"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TxtSizeX 
            Height          =   285
            Left            =   720
            TabIndex        =   138
            Text            =   "2"
            ToolTipText     =   "Horizontal Logosize"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Size X"
            Height          =   255
            Left            =   0
            TabIndex        =   143
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Size Y"
            Height          =   255
            Left            =   0
            TabIndex        =   142
            Top             =   480
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   0
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1500
      Index           =   4
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   8
      Top             =   11640
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.VScrollBar VScrPic 
      Height          =   1575
      LargeChange     =   5
      Left            =   8745
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScrPic 
      Height          =   255
      LargeChange     =   5
      Left            =   0
      TabIndex        =   3
      Top             =   6405
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1500
      Index           =   0
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1500
      Index           =   3
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   6
      Top             =   10080
      Visible         =   0   'False
      Width           =   9000
   End
   Begin ComctlLib.TabStrip TabChoice 
      Height          =   3735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6588
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Foreground"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Border"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Background"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Draw"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmBlock 
      BorderStyle     =   0  'Kein
      Height          =   5040
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9015
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1500
      Index           =   1
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5040
      Width           =   9000
   End
   Begin VB.PictureBox PicTexture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   420
      Index           =   2
      Left            =   3360
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   130
      Top             =   6360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PicTexture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   420
      Index           =   3
      Left            =   2760
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   131
      Top             =   6360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PicTexture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   420
      Index           =   0
      Left            =   600
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PicTexture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   420
      Index           =   1
      Left            =   1200
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label LblLngPack 
      Caption         =   "Scythe"
      Height          =   255
      Left            =   6240
      TabIndex        =   135
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblErrLoad 
      Caption         =   "Number of missing Pictures"
      Height          =   255
      Left            =   4080
      TabIndex        =   134
      ToolTipText     =   "Errors loading Textures"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblWarning 
      Caption         =   "Warning"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Calculating new Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuLoad 
         Caption         =   "Load LMF"
      End
      Begin VB.Menu MnuSaveLMF 
         Caption         =   "Save LMF"
      End
      Begin VB.Menu MnuSavePic 
         Caption         =   "Save Picture"
      End
   End
   Begin VB.Menu MnuLanguage 
      Caption         =   "Language"
      Begin VB.Menu MnuLng 
         Caption         =   "English"
         Checked         =   -1  'True
         HelpContextID   =   1033
         Index           =   0
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "LogoMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'LogoMan V7 by Scythe
'Scythe@cablenet.de

'New in V7
'Some speedups and some errorfixes
'Embos/Engrave for Fonts & Textures
'New design
'Save Logomanfiles in lmf format
' so we can create our own standards
' without clicking every button every time
'Preview the logo in common dialog
'added new Language pack with better functions
'......

'V1 was created after io saw By Ben Jones Textured Text example on PSC

'Some of the Code was
'Optimized By    : NYxZ Software Development and VB Optimizations
'NYxZ on the web : http://nyxz.homepage.dk || nyxz@vip.cybercity.dk

'V6 has a new great function
'Draw Backlight (Sub DrawLight) to create realistic looking iluminated Logos
'Some other new functions and some errorfixing in V6

'This Program uses 4 Picboxes to create the Font
'Pic(0) the Font
'Pic(1) the Visible Picture
'Pic(2) the tiled Texture
'Pic(3) th finished Font
'V4 has his own Finished Font Pic cause we need the font without Background for some functions (Bend Horizontal...)

'There are many Frames
'Only 2 of them are realy importand the rest is only for Optic
'Frame1 (Write Text)
'FrmMainTab(3) (Draw on Picture)

'New Features in LogoMan IV:
'
'Font Position
' Left, Right, Center,  Place the font on any Position
'
'Anti Alias (thx to Kevin Gadd for the Idea from PSC)
'   Make real round fonts
'
'Set Borderposition
'  Dis/Enableable Border left, right, top, bottom
'
'Alpha Blend
'  Create Transparent Fonts
'
'Forground Text Transparent
'  Show only the Border
'
'Negative Picture
' Invert the Complete Logo
'
'Drop real Shadows (Thx to Christian Frerichs for the Idea)
'  Move the Light to crate real Shadows
'
'Multi Language Support
' Create ur own Language Packs
' German & English are included
'
'Mirror Text
' Flip the text Horizontal
'
'Invert (Font)
' Inverts the Background everywhere the Text Hits
'
'Lighten/Darken Texture
' (Thx to visualcode for this idea)
'
'Border Fade (Thx to tibsian from PSC)
' create a border that fades between background and logo
' Its realy hard if u use 2 textures :o)
'
'
'Fixed some Bugs:
'
'Errors thru Typing in Fontname & Fontstyle (thx to RJ Soft of West Tennessee for the Tipp)
'There was an Error in the Borderroutine (thx to tibisan from PSC)
'Changed an "error" in change Text (thx to Saifudheen A A from PSC)
'The special charactors such as ® wont apear after typing now they will
'
' Borders are now Closed not open like before.
'Some other Bugs i cant remember are fixed 2 :o)
'

'Speeded up the Routines and reduced the Size
' Thx to NyXZ from PSC
'
'If u have any Ideas let me know scythe@cablenet.de
'
'To get some new and cool Fonts
'Take a look at fontfile.com



'This API is used for Multilanguage Support
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
Private Const LOCALE_ILANGUAGE = &H1

'Some apis to manipulate pictures
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'For BiBlt (faster than cls)
Private Const WHITENESS = &HFF0062

'Thaks to IStuff for the Help and Info about Font Smothing
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETFONTSMOOTHING = 74
Private Const SPI_SETFONTSMOOTHING = 75
Private Const SPIF_SENDWININICHANGE = &H2

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type


Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0

Dim Re As RECT
Dim Backpic As Boolean
Dim PaintMode As Byte
Dim PaintColor As Long
Dim PaintX, PaintY As Integer
Dim TexPath As String
Dim EmbPath As String
Dim Funkt As Byte
Dim TextX As Long
Dim TextY As Long
Dim BorderSize As Byte
Private Const PI = 3.141592
Dim AppPath As String


Private Type RGBcolor
 R As Long
 G As Long
 B As Long
End Type

Private Type LanguagePack
 Orginal As String
 New     As String
 ToolTip As String
End Type

Private Type POINTAPI
 X As Single
 Y As Single
End Type

'For Save as LMF
Dim TextureName(4) As String
Dim EnableDraw As Boolean

Dim SinCos(3600) As POINTAPI  'Needed for Backlight
Dim ChangedPic As Boolean   'Is set if the User changeg something in Draw Tab

'+----------------------------------------------------------+
'| Frame 1                                                  |
'| This Frame holds all Informations/Buttons u see at Start |
'+----------------------------------------------------------+

'# Font and Text Settings
'# All Comboboxes need to events
'# Change for Keyboard and Click for Mouse


'Change Fontname
Private Sub CboName_Change()
 Pic(0).Font = CboName.Text
 DrawPic
End Sub
Private Sub CboName_Click()
 CboName_Change
End Sub

'Change Fontsize
Private Sub CboSize_Change()
 'Only change the Fontsize if it´s bigger than 0
 If LenB(CboSize) Then
  If Val(CboSize) > 0 And Val(CboSize) < 500 Then
   Pic(0).FontSize = Val(CboSize)
   DrawPic
  End If
 End If
End Sub
Private Sub CboSize_LostFocus()
 'If this lost the focus and there is no size change to 12
 If Val(CboSize) < 1 Then
  CboSize = 12
 End If
End Sub
Private Sub CboSize_Click()
 CboSize_Change
End Sub

'Change Fontstyle
Private Sub LMBStytle_Click(Index As Integer, ButtonDown As Boolean)
 With Pic(0)
 Select Case Index
 Case 0
  Pic(0).FontBold = ButtonDown
 Case 1
  Pic(0).FontItalic = ButtonDown
 Case 2
  Pic(0).FontUnderline = ButtonDown
 End Select
 End With
 DrawPic
End Sub

'Text Position
Private Sub LMBAlign_Click(Index As Integer, ButtonDown As Boolean)
 Dim I As Long

 'All buttons Up
 For I = 0 To 3
  LMBAlign(I).SetButtonstate False
 Next I

 LMBAlign(Index).SetButtonstate True

 With Pic(1)
 'Check if Place is choosen
 If Index = 3 And ButtonDown = True Then
  .Cls
  DrawBackpic
  .Refresh
  'Set Function 3
  'in Pic_Mouse.... Move or Down
  'Function 3 says Place font
  Funkt = 3
  'Set mousepointer to Cross
  .MousePointer = 2
  'Disable the Backlight
  ChkLight.Value = 0
 Else
  'Draw Picture with new settings
  DrawPic
 End If
 End With
End Sub

'Anti Alias
Private Sub ChkAlias_Click()
 DrawPic
End Sub

'Alpha Blending
Private Sub ChkAlpha_Click()
 DrawPic
End Sub

'Embos or Engrave the font
Private Sub ChkEmbosFont_Click()
 If ChkEmbosFont.Value Then
  ChkEngraveFont.Value = 0
 End If
 CheckScroller
End Sub
Private Sub ChkEngraveFont_Click()
 If ChkEngraveFont.Value Then
  ChkEmbosFont.Value = 0
 End If
 CheckScroller
End Sub
Private Sub CheckScroller()
 Dim Tmp As Boolean
 If ChkEngraveFont Or ChkEmbosFont Then Tmp = True
 ChkEmEnBox.Enabled = Tmp
 TxtEmEn.Enabled = Tmp
 VScrEmEn.Enabled = Tmp
 DrawPic
End Sub
Private Sub ChkEmEnBox_Click()
 DrawPic
End Sub
Private Sub TxtEmEn_Change()
 If Val(TxtEmEn) < 8 And Val(TxtEmEn) > 0 Then
  VScrEmEn.Value = -Val(TxtEmEn)
 End If
End Sub
Private Sub VScrEmEn_Change()
 TxtEmEn = Abs(VScrEmEn.Value)
 DrawPic
End Sub


'Tabcontrol
'Holds all frames (buttons,Combos....)
Private Sub TabChoice_Click()
 Dim I As Long

 'Hide all
 For I = 0 To 3
  FrmMainTab(I).Visible = False
 Next I
 'Show what we need to show
 FrmMainTab(TabChoice.SelectedItem.Index - 1).Visible = True
 If TabChoice.SelectedItem.Index = 4 Then
  FrmSettings.Visible = False
 Else
  FrmSettings.Visible = True
 End If
 'Show subframe
 For I = 0 To 10
  If OptTab0(I).Value And OptTab0(I).Enabled = True Then
   FrmSub0(I).Visible = True
  End If
 Next I

 If TabChoice.SelectedItem.Index = 4 Then
  'The Draw Button
  'This button switches between Frame1 and FrmMainTab(3)
  'Change settings for Pic(1)
  'Mousepointer to cross
  Pic(1).MousePointer = 2
  Pic(1).AutoRedraw = False
  'If backlight is on then turn off and set changepic true
  If ChkLight.Value = 1 Then
   ChangedPic = True
   ChkLight.Value = 0
  End If
  'Disable the Shadow function if the user enabled it last time
  ShpSun.Visible = False
 Else
  If ChangedPic = True Then
   If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, CmdRedraw.Caption) <> vbOK Then
    TabChoice.SelectedItem = TabChoice.Tabs(4)
    ChkLight.Value = 0
    Exit Sub
   End If
  End If
  ChangedPic = False
 End If

End Sub

'Optionbuttons to switch between all possible screens
Private Sub OptTab0_Click(Index As Integer)
 Dim I As Long
 Dim Tmp As Boolean

 'Hide all frames
 For I = 0 To 10
  FrmSub0(I).Visible = False
 Next I
 'Show the one we need
 FrmSub0(Index).Visible = True

 'Disable Borderfade if Solid Color is not choosen....
 If Index > 3 And Index < 8 And Index <> 5 Then ChkGlow.Value = 0
 If Index > 8 And Index < 11 Then Backpic = False: CmdClearBack_Click
 If Index > 3 And Index < 7 Then Tmp = True
 ChkTop.Enabled = Tmp
 ChkLeft.Enabled = Tmp
 ChkRight.Enabled = Tmp
 ChkBottom.Enabled = Tmp
 DrawPic
End Sub

'Alpha Blend
Private Sub TxtAlpha_Change()
 If Val(TxtAlpha) < 99 And Val(TxtAlpha) > 0 Then
  'Write the new size to the Scrollbar
  '-Val(TxtAlpha) because the Scrollbar is Negative
  VScrAlpha.Value = -Val(TxtAlpha)
 End If
End Sub
Private Sub VScrAlpha_Change()
 TxtAlpha = Abs(VScrAlpha.Value)
 DrawPic
End Sub

'Border
Private Sub ChkBorder_Click()
 Dim Tmp As Boolean
 Dim I As Long
 'En/Disable all controls on BorderFrame if Activated or not
 If ChkBorder.Value Then
  Tmp = True
  BorderSize = Val(TxtBordersize)
 Else
  BorderSize = 0
 End If
 For I = 4 To 7
  OptTab0(I).Enabled = Tmp
 Next I
 ChkTop.Enabled = Tmp
 ChkLeft.Enabled = Tmp
 ChkBottom.Enabled = Tmp
 ChkRight.Enabled = Tmp
 TxtBordersize.Enabled = Tmp
 VScroll1.Enabled = Tmp
 TabChoice_Click
 DrawPic
End Sub
Private Sub TxtBordersize_Click()
 'U cant set the Bordersize using Keyboard
 VScroll1.SetFocus
End Sub
Private Sub VScroll1_Change()
 'setting Bordersize and show it to the User
 'Abs(VScroll1.Value) because the Value of Vscroll1 is negativ (-2)
 BorderSize = Abs(VScroll1.Value)
 TxtBordersize = BorderSize
 DrawPic
End Sub

'De/Activate MirrorText
Private Sub ChkMirror_Click()
 DrawPic
End Sub

'De/Activate InverterText
Private Sub ChkInvert_Click()
 Dim Tmp As Boolean
 If ChkInvert.Value = 0 Then
  Tmp = True
 End If
 ChkAlpha.Enabled = Tmp
 ChkAlias.Enabled = Tmp
 TxtAlpha.Enabled = Tmp
 VScrAlpha.Enabled = Tmp
 CmdShadow.Enabled = Tmp
 DrawPic
End Sub

'Changing the Text
Private Sub Text1_Change()
 DrawPic
End Sub
Private Sub Text1_LostFocus()
 'If the user Left without any text write "No Text"
 If LenB(Text1) = 0 Then
  Text1 = "No Text"
 End If
End Sub



'+--------+
'| Menus  |
'+--------+

'Load Picture
Private Sub MnuLoad_Click()
 Dim I As Long
 Dim Tmp2 As String
 Dim ers As Long
 Dim FName As String
 Dim FontN As String
 Dim ErrTxt As String

 'Holds the Fileversion
 'Maybe i need this in later versions
 Dim Vers As Byte


 On Error GoTo ErrLoad

 'Me.Tag is set if the Programm was called thru double clicking a lmf file
 If Me.Tag = "" Then
  'Open hooked filedialog
  FName = OpenDialog(Me.hwnd, AppPath & "Saves\", "LogoMan *.lmf" & vbNullChar & "*.lmf" & vbNullChar & vbNullChar, True)
 Else
  'Load the clicked file
  FName = Me.Tag
  Me.Tag = ""
 End If


 If FName <> "" Then
  EnableDraw = False
  Open FName For Binary Access Read As #1

   Get #1, , Vers 'Savefile Version
   Get #1, , I 'Adress for Preview / not needed here

   'We saved Boolean and Checkboxvalues as integer
   'to get them we use GetInteger
   ChkAlias.Value = GetInteger
   ChkMirror.Value = GetInteger
   ChkInvert.Value = GetInteger
   ChkAlpha.Value = GetInteger
   ChkEmbosFont.Value = GetInteger
   ChkEngraveFont.Value = GetInteger
   ChkEmEnBox.Value = GetInteger
   For I = 0 To 10
    OptTab0(I).Value = GetInteger
   Next I
   For I = 0 To 2
    LMBStytle(I).SetButtonstate GetInteger
   Next I
   For I = 0 To 3
    LMBAlign(I).SetButtonstate GetInteger
   Next I
   ChkTex(0).Value = GetInteger
   HScrDL(2).Value = GetInteger
   CboGrStyle(0).ListIndex = GetInteger
   CmdGrCol(0).BackColor = GetColor
   CmdGrCol(1).BackColor = GetColor
   ChkBorder.Value = GetInteger
   ChkLeft.Value = GetInteger
   ChkRight.Value = GetInteger
   ChkTop.Value = GetInteger
   ChkBottom.Value = GetInteger
   ChkTex(1).Value = GetInteger
   CmdSolCol(1).BackColor = GetColor
   ChkGlow.Value = GetInteger
   CboGrStyle(1).ListIndex = GetInteger
   CmdGrCol(2).BackColor = GetColor
   CmdGrCol(3).BackColor = GetColor
   Backpic = GetInteger
   ChkStretch.Value = GetInteger
   ChkShrink.Value = GetInteger
   ChkTile.Value = GetInteger
   VScrMoveBack.Value = GetInteger
   HScrMoveBack.Value = GetInteger
   Pic(1).BackColor = GetColor
   CboGrStyle(2).ListIndex = GetInteger
   CmdGrCol(4).BackColor = GetColor
   CmdGrCol(5).BackColor = GetColor
   CmdSolCol(0).BackColor = GetColor
   HScrDL(3).Value = GetInteger

   'All Textvariables
   TxtSizeX = GetString
   TxtSizeY = GetString
   TxtAlpha = GetString
   TxtEmEn = GetString
   FontN = GetString 'Gets the Fontname
   CboSize.Text = GetString
   TextX = GetString
   TextY = GetString
   Text1 = GetString
   TextureName(0) = GetString
   TextureName(2) = GetString
   TxtBordersize.Text = GetString
   VScroll1.Value = -Val(TxtBordersize)
   BorderSize = Abs(VScroll1.Value)
   TextureName(1) = GetString
   TextureName(4) = GetString
   TextureName(3) = GetString
  Close

  'Check if the Font is installed on this computer
 On Error Resume Next
 CboName = FontN
 On Error GoTo 0
 'If not then add Font to Error Message
 If FontN <> CboName Then
  ErrTxt = "Font: " & FontN & vbCrLf
  ers = ers + 1
 End If

 'Check for all texture and add error message if not found
 For I = 0 To 3
  If TextureName(I) <> vbNullChar Then
   If InStr(1, TextureName(I), "\") = 0 Then
    If I < 2 Then
     Tmp2 = AppPath & "Textures\" & TextureName(I)
    Else
     Tmp2 = AppPath & "Emb\" & TextureName(I)
    End If
   Else
    Tmp2 = TextureName(I)
   End If
   If Dir$(Tmp2) <> "" Then
    LoadTexture I, Tmp2
   Else
    ers = ers + 1
    ErrTxt = ErrTxt & "Picture: " & TextureName(I) & vbCrLf
   End If
  End If
 Next I
 'Check for backgroundimage
 If TextureName(4) <> vbNullChar Then
  If Left$(TextureName(4), 1) = "@" Then
   TextureName(4) = AppPath & Right$(TextureName(4), Len(TextureName(4)) - 1)
  End If
  If Dir$(TextureName(4)) <> "" Then
   Pic(4).Picture = LoadPicture(TextureName(4))
  Else
   ers = ers + 1
   ErrTxt = ErrTxt & TextureName(4)
  End If
 End If

 'Show errors if there any
 If ers > 0 Then MsgBox LblErrLoad & ers & vbCrLf & ErrTxt, vbCritical, LblErrLoad.ToolTipText

 End If
EnableDraw = True
DrawPic
Exit Sub
ErrLoad:
MsgBox "Error", vbCritical, LblErrLoad.ToolTipText
Close
End Sub
Private Function GetString() As String
 Dim I As Long
 Dim TmpStr() As Byte
 Dim TmpBin As Long

 Get #1, , TmpBin
 If TmpBin > 0 Then
  TmpBin = TmpBin - 1
  ReDim TmpStr(TmpBin)
  Get #1, , TmpStr()
  For I = 0 To TmpBin
   GetString = GetString & Chr$(TmpStr(I))
  Next I
 End If
End Function

Private Function GetInteger() As Integer
 Get #1, , GetInteger
End Function
Private Function GetColor() As Long
 Get #1, , GetColor
End Function


Private Sub MnuSaveLMF_Click()
 Dim I As Long
 Dim FName As String

 FName = SaveDialog(Me.hwnd, AppPath & "Saves\", "LogoMan *.lmf" & vbNullChar & "*.lmf" & vbNullChar & vbNullChar)


 If FName <> "" Then
  'first of all cut the Backgroundimagename if needed
  If Left$(TextureName(4), Len(AppPath)) = AppPath Then
   TextureName(4) = "@" & Right$(TextureName(4), Len(TextureName(4)) - Len(AppPath))
  End If
  'Now open the file and write all data
  Open FName For Binary Access Write As #1
   Put #1, , CByte(2) 'Versionsnumber
   Put #1, , I 'Nothing at the moment (will hold Pictureadress later)
   Put #1, , ChkAlias.Value
   Put #1, , ChkMirror.Value
   Put #1, , ChkInvert.Value
   Put #1, , ChkAlpha.Value
   Put #1, , ChkEmbosFont.Value
   Put #1, , ChkEngraveFont.Value
   Put #1, , ChkEmEnBox.Value
   For I = 0 To 10
    Put #1, , OptTab0(I).Value
   Next I
   For I = 0 To 2
    Put #1, , LMBStytle(I).IsButtonDown
   Next I
   For I = 0 To 3
    Put #1, , LMBAlign(I).IsButtonDown
   Next I
   Put #1, , ChkTex(0).Value
   Put #1, , HScrDL(2).Value
   Put #1, , CboGrStyle(0).ListIndex
   Put #1, , CmdGrCol(0).BackColor
   Put #1, , CmdGrCol(1).BackColor
   Put #1, , ChkBorder.Value
   Put #1, , ChkLeft.Value
   Put #1, , ChkRight.Value
   Put #1, , ChkTop.Value
   Put #1, , ChkBottom.Value
   Put #1, , ChkTex(1).Value
   Put #1, , CmdSolCol(1).BackColor
   Put #1, , ChkGlow.Value
   Put #1, , CboGrStyle(1).ListIndex
   Put #1, , CmdGrCol(2).BackColor
   Put #1, , CmdGrCol(3).BackColor
   Put #1, , Backpic
   Put #1, , ChkStretch.Value
   Put #1, , ChkShrink.Value
   Put #1, , ChkTile.Value
   Put #1, , VScrMoveBack.Value
   Put #1, , HScrMoveBack.Value
   Put #1, , Pic(1).BackColor
   Put #1, , CboGrStyle(2).ListIndex
   Put #1, , CmdGrCol(4).BackColor
   Put #1, , CmdGrCol(5).BackColor
   Put #1, , CmdSolCol(0).BackColor
   Put #1, , HScrDL(3).Value

   PutString TxtSizeX
   PutString TxtSizeY
   PutString TxtAlpha
   PutString TxtEmEn
   PutString CboName.Text
   PutString CboSize.Text
   PutString TextX
   PutString TextY
   PutString Text1
   PutString TextureName(0)
   PutString TextureName(2)
   PutString TxtBordersize.Text
   PutString TextureName(1)
   PutString TextureName(4)
   PutString TextureName(3)

   'Save the preview Logo
   Dim Hgt As Long
   Dim Wid As Long
   Dim Buf() As Byte
   'resize if needed
   Wid = Pic(1).ScaleWidth
   Hgt = Pic(1).ScaleHeight
   If Wid > 300 Then
    Hgt = Hgt * 300 / Wid
    Wid = 300
   End If
   If Hgt > 50 Then
    Wid = Wid * 50 / Hgt
    Hgt = 50
   End If
   Pic(0).Width = Wid
   Pic(0).Height = Hgt
   Pic(0).Cls
   'Blit the miniatur
   StretchBlt Pic(0).hdc, 0, 0, Wid, Hgt, Pic(1).hdc, 0, 0, Pic(1).ScaleWidth, Pic(1).ScaleHeight, vbSrcCopy
   'copy to buffer
   Pic2Array1D Pic(0), Buf()
   CompressArray Buf

   'Hold fileposition
   I = Seek(1)
   'write picture data
   Put #1, , Hgt
   Put #1, , Wid
   Put #1, , Buf
   'write the adress of the preview
   Put #1, 2, I
   Pic(0).Width = Pic(1).Width
   Pic(0).Height = Pic(1).Height
  Close
 End If

End Sub
Private Sub PutString(ByVal StringToPut As String)
 Put #1, , CLng(Len(StringToPut))
 Put #1, , StringToPut
End Sub

'Show about dialog
'including Users Windows Language ID
'and creator of the Language Pack
Private Sub MnuAbout_Click()
 MsgBox "LogoMan V7 by Scythe" & vbCrLf & _
 "scythe@cablenet.de" & vbCrLf & vbCrLf & _
 "www.scythe-tools.de" & vbCrLf & vbCrLf & _
 "Your LanguageID = " & GetLanguageID & vbCrLf & vbCrLf & _
 "Language Pack by " & vbCrLf & _
 LblLngPack, vbInformation, _
 "About LogoMan"
End Sub

'Change the language pack
Private Sub MnuLng_Click(Index As Integer)
 Dim I As Long
 On Error GoTo Done
 'Remove all checks
 For I = 0 To 999
  MnuLng(I).Checked = False
 Next I
Done:
 'Now check selected Language
 On Error GoTo 0
 MnuLng(Index).Checked = True
 'Load LngPack
 LoadLanguage Me, AppPath & "LangPacks\" & MnuLng(Index).HelpContextID & ".lng"
 ''Save last used language into lng.ini
 Open AppPath & "LangPacks\lng.ini" For Binary Access Write As #1
  Put #1, , CLng(MnuLng(Index).HelpContextID)
 Close
End Sub

'Save our Logo as BMP
'I cant implement JPG because PSC will kill the jpg.dll
'And "compressed" GIF has a copyright :(
Private Sub MnuSavePic_Click()

 Dim FName As String

 FName = SaveDialog(Me.hwnd, AppPath & "Saves\", "Bitmap *.bmp" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar)
 If FName <> "" Then
  SavePicture Pic(1).Image, FName
 End If
End Sub



'# Texture Settings
'# here we work with Index
'# cause the we have 2 Texture settings
'# 0 for Forground Text and 1 for Border

'Choosen an Texture
Private Sub File_Click(Index As Integer)
 'There is an extra routine for FileLoading
 'Because we need this 2 times
 If Index < 2 Then
  LoadTexture Index, TexPath & File(Index).filename
 Else
  LoadTexture Index, EmbPath & File(Index).filename
 End If
 TextureName(Index) = File(Index).filename
End Sub
Private Sub CmdLoad_Click(Index As Integer)
 Dim FName As String

 FName = OpenDialog(Me.hwnd, AppPath, "Pictures" & vbNullChar & "*.bmp;*.jpg;*.gif;*.wmf;*.emf" & vbNullChar & vbNullChar)

 If FName <> "" Then
  'Set the Filename of the Texturefile into the Tag of the SaveTexture Button
  'and enable it (to Tag so we dont need any new Variable)
  CmdToTex(Index).Tag = FName
  CmdToTex(Index).Enabled = True
  LoadTexture Index, FName
  DrawPic
 End If
End Sub

'Sub routines For Texture Loading
Private Sub LoadTexture(ByVal Index As Integer, ByVal FName As String)
 'The Preview
 ImgPrev(Index).Picture = LoadPicture(FName)
 'The real texture we use for drawings
 PicTexture(Index).Picture = LoadPicture(FName)
 'We use Real White (&HFFFFFF) as Transparent
 'So change this Color in every Texture we load
 ChangeWhite (Index)
 'If Solid Color or Gradient are choosen disable them
 'they will redraw the Picture
 DrawPic
End Sub
Private Sub ChangeWhite(Index As Integer)
 'We need an Rect for Transparent Blitting
 'The Rect tells us how big the part is we want to blit
 Dim rc As RECT
 'Full Picture Size
 With rc
 .Left = 0
 .Top = 0
 .Right = PicTexture(Index).Width
 .Bottom = PicTexture(Index).Height
 End With
 With Pic(2)
 'Pic(2) is not used at the Moment
 'Set new Size for pic(2)
 .Width = PicTexture(Index).Width
 .Height = PicTexture(Index).Height
 'Set Backcolor to &HFFFBFB
 'Then blit picture over it
 'We select color &HFFFFFF as Transparent
 'After this &HFFFFFF is replaced thru &HFFFBFB
 .BackColor = &HFFFBFB
 .Cls
 TransparentBlt .hdc, .hdc, PicTexture(Index).hdc, rc, 0, 0, &HFFFFFF
 BitBlt PicTexture(Index).hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
 'Set Pic(2) Size back to Normal
 .Width = Pic(1).Width
 .Height = Pic(1).Height
 .BackColor = &HFFFFFF
 End With
End Sub

'Copy an Loaded Texture to the Texture Path
Private Sub CmdToTex_Click(Index As Integer)
 'Error Statement if the texture already is there
On Error Resume Next
'Disable Button
CmdToTex(Index).Enabled = False
'Copy File
FileCopy CmdToTex(Index).Tag, TexPath & CutBefore(CmdToTex(Index).Tag, "\")
End Sub

'Darken/Lighten a Texture
'Embos/Engrave Texture
Private Sub HScrDL_Change(Index As Integer)
 If Index > 1 Then
  DrawPic
  Exit Sub
 End If
 Dim Buf() As RGBQUAD
 Dim Buf2() As RGBQUAD
 Dim X As Long
 Dim Y As Long
 Dim Col As Long
 Dim SetIt As Boolean


 'Permanent or preview
 If Index > 1 Then
  Index = Index - 2
  SetIt = True
  CmdDLOK(Index).Enabled = False
 Else
  If HScrDL(Index).Value <> 0 Then CmdDLOK(Index).Enabled = True
 End If

 'Nothing changen (called from DrawPic sub)
 If HScrDL(Index).Value = 0 Then Exit Sub

 'get the Picture
 Pic2Array PicTexture(Index), Buf()

 'create second array
 ReDim Buf2(0 To PicTexture(Index).ScaleWidth - 1, 0 To PicTexture(Index).ScaleHeight)

 'Change the picture
 For X = 0 To PicTexture(Index).ScaleWidth - 1
  For Y = 0 To PicTexture(Index).ScaleHeight - 1
   Col = Buf(X, Y).rgbBlue + HScrDL(Index).Value
   If Col > 255 Then Col = 255
   If Col < 0 Then Col = 0
   Buf2(X, Y).rgbBlue = Col
   Col = Buf(X, Y).rgbGreen + HScrDL(Index).Value
   If Col > 255 Then Col = 255
   If Col < 0 Then Col = 0
   Buf2(X, Y).rgbGreen = Col
   Col = Buf(X, Y).rgbRed + HScrDL(Index).Value
   If Col > 255 Then Col = 255
   If Col < 0 Then Col = 0
   Buf2(X, Y).rgbRed = Col
  Next Y
 Next X

 'Update picture
 Array2Pic PicTexture(Index), Buf2()

 'Only preview mode or update ?
 If SetIt Then
  DrawPic
  HScrDL(Index).Value = 0
 Else
  'Update Preview
  Set ImgPrev(Index).Picture = PicTexture(Index).Image
  'Set old Picture
  Array2Pic PicTexture(Index), Buf()
 End If
End Sub

Private Sub CmdDLOK_Click(Index As Integer)
 HScrDL_Change (Index + 2)
End Sub

'Changed the Gradient Style
'So Activate Gradient
Private Sub CboGrStyle_Change(Index As Integer)
 DrawPic
End Sub
Private Sub CboGrStyle_Click(Index As Integer)
 'TexPath ="" until the Program is complete Loaded
 'We check it so avoid Changes bevore
 If LenB(TexPath) Then
  CboGrStyle_Change (Index)
 End If
End Sub

'Changed the Gradient Color
'So Activate Gradient
Private Sub CmdGrCol_Click(Index As Integer)
 'We have 6 Buttons
 '0 and 2 for Forground
 '1 and 3 for Border
 '6 for Backlight
 Dim Col As Long
 Col = ColorDialog(Me.hwnd, CmdGrCol(Index).BackColor)
 If Col <> -1 Then
  If Col = vbWhite Then Col = &HFFFBFB
  CmdGrCol(Index).BackColor = Col
  DrawPic
 End If
End Sub

'Swap the Gradient Colors
Private Sub CmdSwapColors_Click(Index As Integer)
 Dim X As Long
 With CmdGrCol(Index)
 X = CmdGrCol(Index + 1).BackColor
 CmdGrCol(Index + 1).BackColor = .BackColor
 .BackColor = X
 End With
 DrawPic
End Sub

'Enable Embos/Engrave Textures
Private Sub ChkTex_Click(Index As Integer)
 Dim Tmp As Boolean
 If ChkTex(Index).Value Then
  Tmp = True
 Else
  Tmp = False
  CmdToTex(Index).Enabled = Tmp
 End If
 Index = Index + 2
 CmdLoad(Index).Enabled = Tmp
 File(Index).Enabled = Tmp
 HScrDL(Index).Enabled = Tmp
 DrawPic
End Sub


'Make the Solid Border Glow
Private Sub ChkGlow_Click()
 Dim Tmp As Boolean
 If ChkGlow.Value Then
  Tmp = False
 Else
  Tmp = True
 End If
 ChkTop.Enabled = Tmp
 ChkBottom.Enabled = Tmp
 ChkLeft.Enabled = Tmp
 ChkRight.Enabled = Tmp
 DrawPic
End Sub
Private Sub CmdSolCol_Click(Index As Integer)
 Dim Col As Long
 Col = ColorDialog(Me.hwnd, CmdSolCol(Index).BackColor)
 If Col <> -1 Then
  If Col = vbWhite Then Col = &HFFFBFB
  CmdSolCol(Index).BackColor = Col
  DrawPic
 End If
End Sub


'Background Light
Private Sub ChkLight_Click()
 If ChkLight.Value Then
  'Warn the User that all changes will get lost
  If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, ChkLight.Caption) = 2 Then ChkLight.Value = 0: Exit Sub
  DrawPic
  FrmLight.Enabled = True
 Else
  FrmLight.Enabled = False
 End If
End Sub
Private Sub CmdLightCol_Click()
 Dim Col As Long
 Col = ColorDialog(Me.hwnd, CmdLightCol.BackColor)
 If Col <> -1 Then
  If Col = vbWhite Then Col = &HFFFBFB
  CmdLightCol.BackColor = Col
  DrawPic
 End If
End Sub


'Borders
Private Sub ChkTop_Click()
 DrawPic
End Sub
Private Sub ChkBottom_Click()
 DrawPic
End Sub
Private Sub ChkLeft_Click()
 DrawPic
End Sub
Private Sub ChkRight_Click()
 DrawPic
End Sub


'# Background Settings
'# Here we set Background Pic/Color

'Set the Backcolor for Pic(1)
Private Sub CmdBackcolor_Click()
 Dim Col As Long
 Col = ColorDialog(Me.hwnd, Pic(1).BackColor)
 If Col <> -1 Then
  Pic(1).BackColor = Col
  DrawPic
 End If
End Sub

'Load a Background Picture
Private Sub CmdBackground_Click()
 Dim FName As String

 FName = OpenDialog(Me.hwnd, AppPath, "Pictures" & vbNullChar & "*.bmp;*.jpg;*.gif;*.wmf;*.emf" & vbNullChar & vbNullChar)

 If FName <> "" Then
  Pic(4).Picture = LoadPicture(FName)
  TextureName(4) = FName
  'Set Backpic to true so DrawBackPic know we have one
  Backpic = True
  'End/Disable the Scroll Backpic option
  SetMovePic
  'Enable/Disable the Checkboxes
  If ChkTile.Value = 0 Then
   'Tile is not selected so we can select Stretch/Shrink
   ChkStretch.Enabled = True
   ChkShrink.Enabled = True
  End If
  'Enable Tile Button
  ChkTile.Enabled = True
  'Enable remove Bacground Picture Button
  CmdClearBack.Enabled = True
  DrawPic
 End If
End Sub

'Remove the Bacground Picture
Private Sub CmdClearBack_Click()
 'Tell the DrawBackPic no Picture
 Backpic = False
 Pic(4).Picture = LoadPicture()
 SetMovePic
 'Disable Checkboxes
 ChkStretch.Enabled = False
 ChkShrink.Enabled = False
 ChkTile.Enabled = False
 ChkStretch.Value = 0
 ChkShrink.Value = 0
 ChkTile.Value = 0
 CmdClearBack.Enabled = False
 DrawPic
End Sub

'Stretch the part of our Backgroundpicture thats smaller than our Logo
Private Sub ChkStretch_Click()
 DrawPic
End Sub

'Shrink the part of our Backgroundpicture thats bigger than our Logo
Private Sub ChkShrink_Click()
 DrawPic
 SetMovePic
End Sub

'Tile the Picture
Private Sub ChkTile_Click()
 'If tile is enabled
 'Shrink and Strech must be disabled
 If ChkTile.Value Then
  ChkStretch.Enabled = False
  ChkShrink.Enabled = False
 Else
  ChkStretch.Enabled = True
  ChkShrink.Enabled = True
 End If
 DrawPic
End Sub

'Move The Backgroundpicture
Private Sub HScrMoveBack_Change()
 DrawPic
End Sub
Private Sub VScrMoveBack_Change()
 DrawPic
End Sub
Private Sub SetMovePic()
 'Is the Picture bigger than our Logo ?
 'Let the user Move the background
 If Pic(4).ScaleWidth > Pic(0).ScaleWidth And ChkShrink.Value = 0 And Backpic = True Then
  HScrMoveBack.Max = Pic(4).ScaleWidth - Pic(0).ScaleWidth
  HScrMoveBack.Value = 0
  HScrMoveBack.Enabled = True
 Else
  HScrMoveBack.Enabled = False
 End If
 If Pic(4).ScaleHeight > Pic(0).ScaleHeight And ChkShrink.Value = 0 And Backpic = True Then
  VScrMoveBack.Max = Pic(4).ScaleHeight - Pic(0).ScaleHeight
  VScrMoveBack.Value = 0
  VScrMoveBack.Enabled = True
 Else
  VScrMoveBack.Enabled = False
 End If
End Sub


'# Some Spezials
'# like Copy to Clipboard, Save ....

'Copy our Logo to the Clipboard
'to use it in any HTML designer Paintapp ....
'This button has an index because it´s on Fram1 and Fram5
'We dont need the index but VB does :-)
Private Sub CmdClipboard_Click(Index As Integer)
 'Clear Clipboard from all old
 Clipboard.Clear
 'copy the data of Pic(1) to the Clipboard
 Clipboard.SetData Pic(1).Image
End Sub

'The Size Buttons
'These Textboxes & Scrollbars are in an
'not visible Frame
'we need this because if the frame is not there
'every time u use ur wheel (on Mouse)
'the Picturesize changes
Private Sub VscrX_Change()
 'Change the Horizontal size of our Logo
 TxtSizeX = Abs(VscrX.Value)
 'Check Picture Scorllbars to en/diable
 ChangePicSize
 DrawPic
End Sub
Private Sub TxtSizeX_Change()
 'the minimum Logosize = 20 Pixels
 If Val(TxtSizeX) > 19 And TexPath <> "" Then
  'Write the new size to the Scrollbar
  '-Val(TxtSizeX) because the Scrollbar is Negative
  VscrX.Value = -Val(TxtSizeX)
 End If
End Sub
Private Sub VscrY_Change()
 'Change the Vertical size of our Logo
 TxtSizeY = Abs(VscrY.Value)
 ChangePicSize
 DrawPic
End Sub
Private Sub TxtSizeY_Change()
 If Val(TxtSizeY) > 19 And TexPath <> "" Then
  VscrY.Value = -Val(TxtSizeY)
 End If
End Sub







'+------------------------------------+
'| Buttons for Manupulating the Image |
'+------------------------------------+

'Set one Point in the Picture
Private Sub CmdPoint_Click()
 'Paintmode is used in Pic(1)_Mouse...  Down/Move
 PaintMode = 1
End Sub

'Draw a Line
Private Sub CmdLine_Click()
 PaintMode = 2
 'If PaintX = -1 then we havent set the starting Point of the Line
 PaintX = -1
End Sub

'Draw Filled Box
Private Sub CmdBox_Click()
 PaintMode = 3
 PaintX = -1
End Sub

'Get a color from the Picture
Private Sub CmdPip_Click()
 PaintMode = 4
End Sub

'Select drawcolor from Standard Colordialog
Private Sub CmdDrawColor_Click()
 Dim Col As Long
 Col = ColorDialog(Me.hwnd, CmdDrawColor.BackColor)
 If Col <> -1 Then
  If Col = vbWhite Then Col = &HFFFBFB
  CmdDrawColor.BackColor = Col
  PaintColor = Col
  DrawPic
 End If
End Sub

'Spezial FX (Blur/Monochrome/DePixel)
Private Sub CboFX_Click()
 'Set Mouseicon to Hourglas
 Me.MousePointer = 11

 'We changed the Pic
 ChangedPic = True

 'We use Arrays to Manipulate the Picture
 Dim Buf()        As RGBQUAD    'This Array will hold the RGB Colors from Pic(1)
 Dim X            As Long
 Dim Y            As Long
 Dim I            As Long
 Dim Col          As Long       'Needed to move thru the Array
 Label2.Visible = True          'Show the "Calculating new Image" Label
 Pic(1).Visible = False         'Hide the Picture


 With Pic(1)
 'Fill our Array with the Picture Data
 'Our array moves from bottom to top
 Pic2Array Pic(1), Buf()

 'Check wich FX the user want
 Select Case CboFX.ListIndex
 Case 0 'Monochrome
  'Move thru the Picture
  For X = 0 To .Width - 1
   For Y = 0 To .Height - 1
    'calculate the Colors
    'Red * 0,3 + Green * 0,59 + Blue * 0,11 gives us the Graycolor
    Col = 0.3 * CLng(Buf(X, Y).rgbRed) + 0.59 * CLng(Buf(X, Y).rgbGreen) + 0.11 * CLng(Buf(X, Y).rgbBlue)
    'Write the new Color to all 3 old
    Buf(X, Y).rgbRed = Col
    Buf(X, Y).rgbGreen = Col
    Buf(X, Y).rgbBlue = Col
   Next Y
  Next X
 Case 1 'Blur
  'Move thru the Picture
  For X = 2 To .Width - 2
   For Y = 2 To .Height - 2
    'calculate the new color
    'Add all Colors around our Point (8) and dvide thru 8   (/8)
    'Now give our Pixel the new Color
    'For Red
    Col = (CLng(Buf(X - 1, Y - 1).rgbRed) + CLng(Buf(X - 1, Y).rgbRed) + CLng(Buf(X - 1, Y + 1).rgbRed) + CLng(Buf(X, Y - 1).rgbRed) + CLng(Buf(X, Y + 1).rgbRed) + CLng(Buf(X + 1, Y - 1).rgbRed) + CLng(Buf(X + 1, Y).rgbRed) + CLng(Buf(X + 1, Y + 1).rgbRed)) / 8
    Buf(X, Y).rgbRed = Col
    'For green
    Col = (CLng(Buf(X - 1, Y - 1).rgbGreen) + CLng(Buf(X - 1, Y).rgbGreen) + CLng(Buf(X - 1, Y + 1).rgbGreen) + CLng(Buf(X, Y - 1).rgbGreen) + CLng(Buf(X, Y + 1).rgbGreen) + CLng(Buf(X + 1, Y - 1).rgbGreen) + CLng(Buf(X + 1, Y).rgbGreen) + CLng(Buf(X + 1, Y + 1).rgbGreen)) / 8
    Buf(X, Y).rgbGreen = Col
    'For blue
    Col = (CLng(Buf(X - 1, Y - 1).rgbBlue) + CLng(Buf(X - 1, Y).rgbBlue) + CLng(Buf(X - 1, Y + 1).rgbBlue) + CLng(Buf(X, Y - 1).rgbBlue) + CLng(Buf(X, Y + 1).rgbBlue) + CLng(Buf(X + 1, Y - 1).rgbBlue) + CLng(Buf(X + 1, Y).rgbBlue) + CLng(Buf(X + 1, Y + 1).rgbBlue)) / 8
    Buf(X, Y).rgbBlue = Col
   Next Y
  Next X
 Case 2 'DePixel
  'Move thru the Picture
  For X = 2 To .Width - 2
   For Y = 2 To .Height - 2
    'check if our Pixel is the only of its kind arround (1 Black Pixel sorounded from white)
    'If yes then delete him
    If CLng(Buf(X - 1, Y - 1).rgbRed) = CLng(Buf(X - 1, Y - 1).rgbRed) And CLng(Buf(X - 1, Y).rgbRed) = CLng(Buf(X - 1, Y - 1).rgbRed) And CLng(Buf(X - 1, Y + 1).rgbRed) And CLng(Buf(X, Y - 1).rgbRed) And CLng(Buf(X, Y + 1).rgbRed) And CLng(Buf(X + 1, Y - 1).rgbRed) And CLng(Buf(X + 1, Y).rgbRed) And CLng(Buf(X + 1, Y + 1).rgbRed) = CLng(Buf(X - 1, Y - 1).rgbRed) Then
     Buf(X, Y).rgbRed = Buf(X - 1, Y - 1).rgbRed
    End If
    If CLng(Buf(X - 1, Y - 1).rgbGreen) = CLng(Buf(X - 1, Y - 1).rgbGreen) And CLng(Buf(X - 1, Y).rgbGreen) = CLng(Buf(X - 1, Y - 1).rgbGreen) And CLng(Buf(X - 1, Y + 1).rgbGreen) And CLng(Buf(X, Y - 1).rgbGreen) And CLng(Buf(X, Y + 1).rgbGreen) And CLng(Buf(X + 1, Y - 1).rgbGreen) And CLng(Buf(X + 1, Y).rgbGreen) And CLng(Buf(X + 1, Y + 1).rgbGreen) = CLng(Buf(X - 1, Y - 1).rgbGreen) Then
     Buf(X, Y).rgbGreen = Buf(X - 1, Y - 1).rgbGreen
    End If
    If CLng(Buf(X - 1, Y - 1).rgbBlue) = CLng(Buf(X - 1, Y - 1).rgbBlue) And CLng(Buf(X - 1, Y).rgbBlue) = CLng(Buf(X - 1, Y - 1).rgbBlue) And CLng(Buf(X - 1, Y + 1).rgbBlue) And CLng(Buf(X, Y - 1).rgbBlue) And CLng(Buf(X, Y + 1).rgbBlue) And CLng(Buf(X + 1, Y - 1).rgbBlue) And CLng(Buf(X + 1, Y).rgbBlue) And CLng(Buf(X + 1, Y + 1).rgbBlue) = CLng(Buf(X - 1, Y - 1).rgbBlue) Then
     Buf(X, Y).rgbBlue = Buf(X - 1, Y - 1).rgbBlue
    End If
   Next Y
  Next X
 Case 3 'Negative
  For X = 0 To .Width - 1
   For Y = 0 To .Height - 1
    Buf(X, Y).rgbBlue = 255 - CLng(Buf(X, Y).rgbBlue)
    Buf(X, Y).rgbGreen = 255 - CLng(Buf(X, Y).rgbGreen)
    Buf(X, Y).rgbRed = 255 - CLng(Buf(X, Y).rgbRed)
   Next Y
  Next X
 End Select

 'now update the Picture with our changed data
 Array2Pic Pic(1), Buf()

 'Show Pic and hide Label
 .Visible = True
 Label2.Visible = False
 'Clear Combobox
 CboFX.ListIndex = -1
 'Show Mouseicon
 Me.MousePointer = 0
 End With
End Sub

'Redraw the Orginal Pic
Private Sub CmdRedraw_Click()
 If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, CmdRedraw.Caption) = vbOK Then
  ChangedPic = False
  DrawPic
 End If
End Sub

'Bend (Horizontal = 0 Vertical = 1)
Private Sub CmdBend_Click(Index As Integer)
 'Warn the User that all changes will get lost
 If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, CmdBend(Index).Caption) = 2 Then Exit Sub
 'He pressed OK
 'Clear pic(1) and draw the Background on it
 With Pic(1)
 .AutoRedraw = True
 BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
 DrawBackpic
 .AutoRedraw = False
 'Set the Funkt Variable for Mouse_... Move/Down
 Funkt = Index + 1
 End With
End Sub

'Shadow
Private Sub CmdShadow_Click()
 'Warning that all get lost
 If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, CmdShadow.Caption) = 2 Then Exit Sub
 'Show the Sun on Magnify Picture
 ShpSun.Visible = True
 ShpSun.Tag = "" 'Clean Tag say´s Shadow not hover
 'PicMagnify.Cls
 BitBlt PicMagnify.hdc, 0, 0, PicMagnify.ScaleWidth, PicMagnify.ScaleHeight, 0, 0, 0, WHITENESS
End Sub

'Hover Font
'Thats what windows call shadow :o)
Private Sub CmdHover_Click()
 'Warning that all get lost
 If MsgBox(LblWarning.Caption, vbInformation + vbDefaultButton2 + vbOKCancel, CmdShadow.Caption) = 2 Then Exit Sub
 'set the Tag of ShpSun so we can deside between Shadow and Hover
 ShpSun.Tag = "Hover"
 'Show the Sun on Magnify Picture
 ShpSun.Visible = True
 'PicMagnify.Cls
 BitBlt PicMagnify.hdc, 0, 0, PicMagnify.ScaleWidth, PicMagnify.ScaleHeight, 0, 0, 0, WHITENESS
End Sub

Private Sub PicMagnify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'If the Sun is vissible then move it to the Mouse Coordinates and dra the preview
 If ShpSun.Visible = True Then
  ShpSun.Left = X - ShpSun.Width / 2
  ShpSun.Top = Y - ShpSun.Height / 2
  If ShpSun.Tag = "" Then
   DrawShadow
  Else
   DrawHover
  End If
 Else
  ShpSun.Tag = ""
 End If
End Sub
Private Sub PicMagnify_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Mouse Down now draw the Shadow
 If ShpSun.Visible Then
  If ShpSun.Tag = "" Then
   DrawShadow True
  Else
   DrawHover True
  End If
 End If
 'Disable the sun
 ShpSun.Visible = False
 'PicMagnify.Cls
 BitBlt PicMagnify.hdc, 0, 0, PicMagnify.ScaleWidth, PicMagnify.ScaleHeight, 0, 0, 0, WHITENESS
 'We changed the Pic
 ChangedPic = True
End Sub





'+----------------------------+
'| Form Events                |
'| Used at Start, End, Resize |
'+----------------------------+

'The program Starts
Private Sub Form_Load()


 Dim I As Double
 Dim F As Long
 Dim h As Long
 Dim a As String
 Dim Tmp As String

 'Set the TextSize Textboxes/Scrollbars
 TxtSizeX = Pic(1).ScaleWidth
 TxtSizeY = Pic(1).ScaleHeight
 VscrX.Value = -Val(TxtSizeX)
 VscrY.Value = -Val(TxtSizeY)

 'Get Fontnames and set them to the Font select Box
 For I = 0 To Screen.FontCount - 1
  CboName.AddItem Screen.Fonts(I)
 Next

 'Set the standard Fontsizes
 For I = 12 To 72 Step 4
  CboSize.AddItem I
 Next

 'Now activate the Settings
 CboName.ListIndex = 0        'First Font
 CboSize = 36                 'Fontsize = 36
 CboGrStyle(0).ListIndex = 0  'Font Gradientstyle = Circular
 CboGrStyle(1).ListIndex = 0  'Border Gradientstyle = Circular
 CboGrStyle(2).ListIndex = 0  'Border Gradientstyle = Circular
 LMBAlign(1).SetButtonstate True

 AppPath = App.Path
 'If the last letter of App.path ist not \ then add \
 If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
 'Add "textures\"
 'Now our Texturepath sould look like this "applicationpath\textures\"
 TexPath = AppPath & "textures\"
 'And new in this Version the Embos Picture
 EmbPath = AppPath & "emb\"

 'From now on we need an error handler
 'If the user has the Texturepath not in the Programm Dir
 On Error GoTo NoTexture
 'Add texpath to our 2 Fileboxes
 File(0).Path = TexPath
 File(1).Path = TexPath
 'Now the Embos
 File(2).Path = EmbPath
 File(3).Path = EmbPath

 'Now load the first texture in both Preview Images
 a = TexPath & File(0).List(0)
 LoadTexture 0, a
 LoadTexture 1, a
 a = EmbPath & File(2).List(0)
 LoadTexture 2, a

 'From nowon Logoman shows drawings
 EnableDraw = True

 LoadTexture 3, a
 TextureName(0) = File(0).List(0)
 TextureName(1) = File(0).List(0)
 TextureName(2) = File(2).List(0)
 TextureName(3) = File(2).List(0)

 'Show all Possible Languages
 I = 0
 a = Dir$(AppPath & "LangPacks\*.lng")
 Do Until a = ""
  h = Val(Left$(a, Len(a) - 3))
  Tmp = Trim(GetLanguageName(h))
  If Tmp <> "English" Then
   I = I + 1
   Load MnuLng(I)
   MnuLng(I).Caption = Tmp
   MnuLng(I).Visible = True
   MnuLng(I).HelpContextID = h
   MnuLng(I).Checked = False
  End If
  a = Dir$
 Loop


 'Load Language Pack
 h = 0
 'Check if ther is an Language ini
 If Dir$(AppPath & "LangPacks\lng.ini") = "" Then
  If Dir$(AppPath & "LangPacks\" & Trim(Str(GetLanguageID)) & ".lng") <> "" Then
   h = GetLanguageID
  End If
  'So we start this app the first time
  'Lets create an File Assotiation for .lmf
  'if we are not running ide
  If (App.LogMode = 1) Then
   FileExt ".lmf", AppPath & "LogoMan.exe", "LogoMan Saves", AppPath & "LogoMan.exe,1"
  End If
 Else 'No Language ini so load the userlanguage if possible
  Open AppPath & "LangPacks\lng.ini" For Binary Access Read As #1
   Get #1, , h
  Close
 End If

 If h <> 0 Then
  For F = 0 To I
   If MnuLng(F).HelpContextID = h Then
    MnuLng_Click (F)
    Exit For
   End If
  Next F
 End If

 'Create the Sinus/Cosinus Table
 h = 0
 For I = 0 To 360 Step 0.1
  SinCos(h).X = Cos(I * PI / 180)
  SinCos(h).Y = Sin(I * PI / 180)
  h = h + 1
 Next I

 'Startet file using extension
 If Command <> "" Then Me.Tag = Command: MnuLoad_Click

 Exit Sub
NoTexture:
 MsgBox "Texturepath not found" & vbCrLf & TexPath & vbCrLf & EmbPath, vbCritical, "Error loading Textures"
 End
End Sub

'Pressed the Close Button (X)
Private Sub Form_Terminate()
 'Delete the Program
 Unload Me
 End
End Sub

'Resize the form
Private Sub Form_Resize()
 'Form is Minimized
 With Me
 If .WindowState = 1 Or FrmMainTab(3).Visible = True Then
  'Set the tag of our Form to 1 and exit
  .Tag = 1
  Exit Sub
 End If

 'Never make it to small
 If .ScaleWidth < 600 Then
  .Width = 608 * Screen.TwipsPerPixelX
 End If
 If .ScaleHeight < 380 Then
  .Height = 464 * Screen.TwipsPerPixelY
 End If
 FrmBlock.Width = Me.ScaleWidth

 ChangePicSize
 End With
End Sub






'+-------------+
'| Subroutines |
'+-------------+

'Draw the Logo
'This is #1
Private Sub DrawPic()

 'Form not loaded or Lmf file loading not finished
 If EnableDraw = False Then Exit Sub

 Dim retval As Long         'Needed for Fontsmothing
 Dim Fs As Long             'Needed for Fontsmothing
 Dim X As Long
 Dim Y As Long
 Dim j As Long
 Dim k As Long
 Dim I As Long
 Dim Tmp As Long
 Dim Bor(3) As Long
 Dim Buf() As RGBQUAD
 Dim Buf2() As RGBQUAD
 Dim buf3() As RGBQUAD
 Dim buf4() As RGBQUAD
 Dim Col As Long
 Dim Alp As Long

 'Disable Font Smoothing if enabled
 retval = SystemParametersInfo(SPI_GETFONTSMOOTHING, 0, Fs, 0)
 retval = SystemParametersInfo(SPI_SETFONTSMOOTHING, 0, 0&, SPIF_SENDWININICHANGE)

 'Someone wantent to lighten/darken the textures
 'But he didnt say ok so get back the real preview
 For X = 0 To 1
  If CmdDLOK(X).Enabled = True Then
   Set ImgPrev(X).Picture = PicTexture(X)
   HScrDL(X).Value = 0
   CmdDLOK(X).Enabled = False
  End If
 Next X


 Pic(1).AutoRedraw = True
 'Get the Textposition
 GetTextPos

 'Clear all Pictures
 For X = 0 To 3
  BitBlt Pic(X).hdc, 0, 0, Pic(X).ScaleWidth, Pic(X).ScaleHeight, 0, 0, 0, WHITENESS
 Next X
 Pic(1).Cls

 'Nearly all Operations are with Pic(0)
 With Pic(0)
 'Draw the Background Picture
 DrawBackpic

 X = Int((.ScaleWidth - .TextWidth(Text1)) / 2)
 Y = Int((.ScaleHeight - .TextHeight(Text1)) / 2)
 'Check if Border is enabled
 If ChkLeft.Value Then
  Bor(0) = X - BorderSize
 Else
  Bor(0) = X
 End If
 If ChkRight.Value Then
  Bor(1) = X + BorderSize
 Else
  Bor(1) = X
 End If
 If ChkTop.Value Then
  Bor(2) = Y - BorderSize
 Else
  Bor(2) = Y
 End If
 If ChkBottom.Value Then
  Bor(3) = Y + BorderSize
 Else
  Bor(3) = Y
 End If
 'Draw Border if enabled and FadeBorder is Disabled
 If ChkBorder.Value And OptTab0(7).Value = False Then
  'Set Paintcolor for text
  'Color 1 will be transparent
  If OptTab0(5).Value Then
   'Solid Color
   .ForeColor = CmdSolCol(1).BackColor
  Else
   'Transparent
   .ForeColor = 1
   'If Transparent then check for Gradient or Texture
   If OptTab0(6).Value = False Then
    'Create Texture
    CreateTexture (1)
   Else
    'Create Gradient
    DrawGradient (1)
   End If
  End If





  'Draw Border
  For j = Bor(0) To Bor(1)
   For k = Bor(2) To Bor(3)
    .CurrentX = j
    .CurrentY = k
    Pic(0).Print Text1
   Next k
  Next j

  'If the Foreground Text is Transparent then Draw it now
  If OptTab0(3).Value Then
   .CurrentX = X
   .CurrentY = Y
   .ForeColor = &HFFFFFF
   Pic(0).Print Text1
  End If
  'Now the Border is done
  'Lets mix it
  'Blit the text to pic(2) (here is our Tiled Texture or the Gradient)
  'Normaly (No Solid Color Font) we blit White over the Texture
  'And leave the Black Font empty


  TransparentBlt Pic(2).hdc, Pic(2).hdc, .hdc, Re, 0, 0, 1
  'Now from 2 to 3 with white as Transparent only the Color thats over from 1st blit will be bilted
  TransparentBlt Pic(3).hdc, Pic(3).hdc, Pic(2).hdc, Re, 0, 0, &HFFFFFF
  'Clear pic(0) for normal Text Operation
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
 End If

 '# Draw the Normal Text
 'Set the Color
 If OptTab0(3).Value = False Then
  If OptTab0(1).Value Then
   If OptTab0(1).Value Then
    'Solid Color
    .ForeColor = CmdSolCol(0).BackColor
   Else
    .ForeColor = vbWhite
   End If
  Else
   If OptTab0(2).Value = True Then
    DrawGradient (0)
   Else
    CreateTexture (0)
   End If
   .ForeColor = 1
  End If





  '# If Fade Border is Active then draw it Now
  If OptTab0(7).Value Or ChkGlow.Value Then
   'Store Picture Color
   Tmp = Pic(0).ForeColor
   Bor(0) = BorderSize
   'Now draw the Bordertext
   'Change the Color from out to inner
   'As Color we take the transparency
   For I = 1 To BorderSize
    .ForeColor = Int((100 / (BorderSize + 1 + ChkGlow.Value)) * I) 'Get the Fade Status
    For j = X - Bor(0) To X + Bor(0)
     For k = Y - Bor(0) To Y + Bor(0)
      .CurrentX = j
      .CurrentY = k
      Pic(0).Print Text1
     Next k
    Next j
    Bor(0) = Bor(0) - 1
   Next I
   'Print the text one last Time with color &Hff
   'Every color over 100 wont be checked later
   .CurrentX = X
   .CurrentY = Y
   'Print Text
   .ForeColor = &HFF
   Pic(0).Print Text1
   .ForeColor = Tmp

   'Is this Glow Effect
   If ChkGlow.Value Then
    Pic2Array Pic(2), buf4()
   End If

   'If we use solid Color
   If OptTab0(1).Value Or ChkGlow.Value Then
    Pic(2).BackColor = CmdSolCol(1).BackColor
    Pic(2).Cls
   End If

   'Now get the Pictures we need for checks/calculations
   Pic2Array Pic(1), Buf()   'Texture Back
   Pic2Array Pic(2), Buf2()  'Texture Font
   Pic2Array Pic(0), buf3()  'Font

   'Now set Backcolor to white if we changed bevore
   If OptTab0(1).Value Or ChkGlow.Value Then
    Pic(2).BackColor = &HFFFFFF
    BitBlt Pic(2).hdc, 0, 0, Pic(2).ScaleWidth, Pic(2).ScaleHeight, 0, 0, 0, WHITENESS
    If ChkGlow.Value Then
     Array2Pic Pic(2), buf4()
    End If
   End If

   'Move thru Pic(0) where the Multicolored Font is
   'If there is a Color under 100 then
   'make alphablend between TextTexture & Background
   'Put this Alphablend to the array we have the Font
   For j = 1 To .Width - 1
    For k = 1 To .Height - 1
     If buf3(j, k).rgbRed < 100 Then
      Alp = buf3(j, k).rgbRed
      Col = Int((CLng(Buf(j, k).rgbBlue * (100 - Alp) + CLng(Buf2(j, k).rgbBlue) * Alp)) / 100)
      If Col > 255 Then Col = 255
      buf3(j, k).rgbBlue = Col
      Col = Int((CLng(Buf(j, k).rgbGreen * (100 - Alp) + CLng(Buf2(j, k).rgbGreen) * Alp)) / 100)
      If Col > 255 Then Col = 255
      buf3(j, k).rgbGreen = Col
      Col = Int((CLng(Buf(j, k).rgbRed * (100 - Alp) + CLng(Buf2(j, k).rgbRed) * Alp)) / 100)
      If Col > 255 Then Col = 255
      buf3(j, k).rgbRed = Col
     End If
    Next k
   Next j

   'Now set the Picture we created [old pic(0)] to Pic(3)
   Array2Pic Pic(3), buf3()
   BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
   '.Cls
  End If

  'Set Textposition
  .CurrentX = X
  .CurrentY = Y

  'Print Text
  Pic(0).Print Text1
 End If
 'Blit the Font
 TransparentBlt Pic(2).hdc, Pic(2).hdc, .hdc, Re, 0, 0, 1
 If ChkAlias.Value = 1 And BorderSize <> 0 Then
  TranBltAlias Pic(2), X - BorderSize, Y - BorderSize, Pic(0).TextWidth(Text1) + BorderSize * 2, Pic(0).TextHeight(Text1) + BorderSize * 2, Pic(3), X - BorderSize, Y - BorderSize, &HFFFFFF, True, 100
 Else
  TransparentBlt Pic(3).hdc, Pic(3).hdc, Pic(2).hdc, Re, 0, 0, &HFFFFFF
 End If

 If ChkEmbosFont.Value Or ChkEngraveFont.Value Then
  I = Val(TxtEmEn) + 1
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  If ChkEmEnBox.Value Then
   Dim x1 As Long
   Dim x2 As Long
   Dim y1 As Long
   Dim y2 As Long
   x1 = TextX - I
   y1 = TextY - I
   x2 = TextX + Pic(0).TextWidth(Text1) + BorderSize * 2 + I
   y2 = TextY + Pic(0).TextHeight(Text1) + BorderSize * 2 + I
   Pic(0).Line (x1, y1)-(x2, y2), &HA0A0A0, BF
   Pic(0).Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), 0, BF
   EmbosEngrave 5, x1 - 1, y1, x2, y2 + 1
  Else
   Tmp = .ForeColor
   .ForeColor = &HA0A0A0
   For j = TextX - I To TextX + BorderSize * 2 + 1 + I
    For k = TextY - I To TextY + BorderSize * 2 + 1 + I
     .CurrentX = j
     .CurrentY = k
     Pic(0).Print Text1
    Next k
   Next j
   .ForeColor = 0
   For j = TextX - I + 1 To TextX + BorderSize * 2 + I
    For k = TextY - I + 1 To TextY + BorderSize * 2 + I
     .CurrentX = j
     .CurrentY = k
     Pic(0).Print Text1
    Next k
   Next j
   .ForeColor = Tmp
   EmbosEngrave 5, TextX - BorderSize - I, TextY - BorderSize - I, TextX + Pic(0).TextWidth(Text1) + BorderSize * 2 + I, TextY + Pic(0).TextHeight(Text1) + BorderSize * 2 + I
  End If
  BitBlt Pic(0).hdc, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleHeight, 0, 0, 0, WHITENESS
 End If


 'Set the Rect for the last Copyroutine
 'Copy the finished Text to pic(1)
 Dim rc As RECT
 rc.Left = X - BorderSize
 rc.Right = rc.Left + .TextWidth(Text1) + BorderSize * 2
 rc.Top = Y - BorderSize
 rc.Bottom = rc.Top + .TextHeight(Text1) + BorderSize * 2
 End With
 If rc.Left < 0 Then rc.Left = 0
 If rc.Right > TxtSizeX Then rc.Right = TxtSizeX
 If rc.Top < 0 Then rc.Top = 0
 If rc.Bottom > TxtSizeY Then rc.Bottom = TxtSizeY
 If TextX < 0 Then TextX = 0
 If TextY < 0 Then TextY = 0

 'Check Alpha Blending
 If ChkAlpha.Value = 1 Then
  X = Val(TxtAlpha)
 Else
  X = 100
 End If

 'Mirror the whole Text
 If ChkMirror Then
  StretchBlt Pic(3).hdc, Pic(3).Width, 0, -Pic(3).Width, Pic(3).Height, Pic(3).hdc, 0, 0, Pic(3).Width, Pic(3).Height, vbSrcCopy
 End If

 If ChkInvert Then
  BltInvert Pic(3), rc.Left, rc.Top, Pic(0).TextWidth(Text1) + BorderSize * 2, Pic(0).TextHeight(Text1) + BorderSize * 2, Pic(1), TextX, TextY, &HFFFFFF
 Else
  'New Check if Anti Alias is selectet
  If ChkAlias.Value = 1 Or ChkAlpha.Value = 1 Then
   Dim Tmp1 As Boolean
   If ChkAlias.Value = 1 Then
    Tmp1 = True
   End If
   Y = Pic(0).TextHeight(Text1)
   TranBltAlias Pic(3), rc.Left, rc.Top, Pic(0).TextWidth(Text1) + BorderSize * 2, Pic(0).TextHeight(Text1) + BorderSize * 2, Pic(1), TextX, TextY, &HFFFFFF, Tmp1, X
  Else
   TransparentBlt Pic(1).hdc, Pic(1).hdc, Pic(3).hdc, rc, TextX, TextY, &HFFFFFF
  End If
 End If

 'Set Fontsmoothing to its Normal State
 retval = SystemParametersInfo(SPI_SETFONTSMOOTHING, 0, Fs, SPIF_SENDWININICHANGE)
End Sub

'Draw Background Picture
Private Sub DrawBackpic()
 'Check if there is an Background ?
 If Backpic = False And OptTab0(10).Value = False Then Exit Sub

 If OptTab0(10).Value Then
  Pic(2).BackColor = Pic(1).BackColor
  DrawGradient 2
  BitBlt Pic(1).hdc, 0, 0, Pic(1).Width, Pic(1).Height, Pic(2).hdc, 0, 0, vbSrcCopy
  Pic(2).BackColor = &HFFFFFF
  Exit Sub
 End If

 Dim X As Long
 Dim Y As Long
 Dim j As Long
 Dim k As Long
 With Pic(1)
 'Get size of the Picture we see
 Y = .Height
 X = .Width

 'Tile the Picture ?
 If ChkTile.Value Then
  'Move thru the pic we see as often as our Background fit´s in
  For Y = 0 To .Height Step Pic(4).Height
   For X = 0 To .Width Step Pic(4).Width
    BitBlt .hdc, X, Y, Pic(4).Width, Pic(4).Height, Pic(4).hdc, 0, 0, vbSrcCopy
   Next X
  Next Y
  'If Stretch  or Shrink is enabled skip all
  ElseIf ChkStretch.Value = 1 Or ChkShrink.Value Then
  'If Stretch is enabled
  If ChkStretch.Value Then
   'Is the Background.Width smaller than the pic we see
   If X > Pic(4).Width Then
    'Get the size of the BG
    j = Pic(4).Width
   Else
    'Use Orginal Size
    j = X
   End If
   'Same for Height
   If Y > Pic(4).Height Then
    k = Pic(4).Height
   Else
    k = Y
   End If
  End If
  'Now with shrink
  If ChkShrink.Value Then
   'If the BG is bigger then shrink
   If .Width < Pic(4).Width Then
    j = Pic(4).Width
   Else
    'J is not set to stretch so hold the org. size
    If j = 0 Then
     j = Pic(4).Width
     X = j
    End If
   End If
   If .Height < Pic(4).Height Then
    k = Pic(4).Height
   Else
    If k = 0 Then
     k = Pic(4).Height
     Y = k
    End If
   End If
  End If
  'Now Blit the BG
  StretchBlt .hdc, 0, 0, X, Y, Pic(4).hdc, 0, 0, j, k, vbSrcCopy
 Else
  'Nothing set so blit only the orginam size
  BitBlt .hdc, 0, 0, .Width, .Height, Pic(4).hdc, HScrMoveBack.Value, VScrMoveBack.Value, vbSrcCopy
 End If
 End With
End Sub

'Tile the Texture over pic(2)
Private Sub CreateTexture(PicIndex As Integer)
 Dim X As Long
 Dim Y As Long
 'get size of the Texture
 With Pic(2)
 'Move thru the pic(2) as often as our texture fit´s in
 For Y = 0 To .Height Step PicTexture(PicIndex).Height
  For X = 0 To .Width Step PicTexture(PicIndex).Width
   BitBlt .hdc, X, Y, PicTexture(PicIndex).Width, PicTexture(PicIndex).Height, PicTexture(PicIndex).hdc, 0, 0, vbSrcCopy
  Next
 Next
 End With
 If ChkTex(PicIndex).Value = 1 Then EmbosEngrave PicIndex, 0, 1, Pic(PicIndex).ScaleWidth, Pic(PicIndex).ScaleHeight
End Sub
'Embos/Engrave the Picture
Private Sub EmbosEngrave(ByVal PicIndex As Long, ByVal StartX As Long, ByVal StartY As Long, ByVal EndX As Long, ByVal EndY As Long)


 Dim X As Long
 Dim Y As Long
 Dim Buf() As RGBQUAD
 Dim Buf2() As RGBQUAD
 Dim Buf1() As RGBQUAD
 Dim R As Long
 Dim G As Long
 Dim B As Long
 Dim NewCol As Single
 Dim Multi As Single
 Dim pos1 As Long
 Dim pos2 As Long
 Dim LookUp(-255 To 255) As Long


 X = StartY
 StartY = Pic(0).ScaleHeight - EndY
 EndY = Pic(0).ScaleHeight - X




 'We Embos/Engrave the Texture
 'If not we work on the Background (Fontembos/grave)
 If PicIndex < 2 Then
  PicIndex = PicIndex + 2
  'Save Pic(3)
  Pic2Array Pic(3), Buf1()
  'Create tiled image from texture
  With Pic(3)
  For Y = 0 To .Height Step PicTexture(PicIndex).Height
   For X = 0 To .Width Step PicTexture(PicIndex).Width
    BitBlt .hdc, X, Y, PicTexture(PicIndex).Width, PicTexture(PicIndex).Height, PicTexture(PicIndex).hdc, 0, 0, vbSrcCopy
   Next
  Next
  End With
  'Embos or Engrave ?
  If HScrDL(PicIndex).Value > 0 Then
   pos1 = 1
   pos2 = 0
  Else
   pos1 = 0
   pos2 = 1
  End If
  'Get pictures for embos
  Pic2Array Pic(2), Buf2()
  Pic2Array Pic(3), Buf()
  Multi = Abs(HScrDL(PicIndex).Value) / 100
 Else
  'We embos the font
  If ChkMirror Then
   StretchBlt Pic(0).hdc, Pic(0).Width, 0, -Pic(0).Width, Pic(0).Height, Pic(0).hdc, 0, 0, Pic(0).Width, Pic(0).Height, vbSrcCopy
  End If
  Pic2Array Pic(1), Buf2()
  Pic2Array Pic(0), Buf()
  Multi = 0.5
  If ChkEngraveFont.Value Then
   pos1 = 1
   pos2 = 0
  Else
   pos1 = 0
   pos2 = 1
  End If
 End If

 'Check if the coordinats are ok
 'if not then fix it
 If StartX < 1 Then StartX = 1
 If StartY < 1 Then StartY = 1
 If EndX > UBound(Buf, 1) - 1 Then EndX = UBound(Buf, 1) - 1
 If EndY > UBound(Buf, 2) - 1 Then EndY = UBound(Buf, 2) - 1

 'Create my all time loved LookUp for faster Math
 For X = -255 To 255
  LookUp(X) = Abs(X * Multi + 128)
 Next X

 'Embos/Engrave
 For Y = StartY To EndY
  For X = StartX To EndX
   R = LookUp(CLng(Buf(X + pos1, Y - pos1).rgbRed) - Buf(X + pos2, Y - pos2).rgbRed)
   G = LookUp(CLng(Buf(X + pos1, Y - pos1).rgbGreen) - Buf(X + pos2, Y - pos2).rgbGreen)
   B = LookUp(CLng(Buf(X + pos1, Y - pos1).rgbBlue) - Buf(X + pos2, Y - pos2).rgbBlue)
   NewCol = (B + G + R) / 384

   R = NewCol * Buf2(X, Y).rgbRed
   G = NewCol * Buf2(X, Y).rgbGreen
   B = NewCol * Buf2(X, Y).rgbBlue
   If R > 255 Then R = 255
   If G > 255 Then G = 255
   If B > 255 Then B = 255

   Buf2(X, Y).rgbRed = CLng(R)
   Buf2(X, Y).rgbGreen = CLng(G)
   Buf2(X, Y).rgbBlue = CLng(B)
  Next X
 Next Y

 'Show the results
 If PicIndex > 4 Then
  Array2Pic Pic(1), Buf2()
 Else
  Array2Pic Pic(2), Buf2()
  Array2Pic Pic(3), Buf1() 'restore pic(3)
 End If

End Sub

'Get the actual Textposition
Private Sub GetTextPos()
 Dim X As Integer
 Dim Y As Integer
 With Pic(0)
 'Set x & y Position to center (case 0)
 X = Int((.ScaleWidth - .TextWidth(Text1)) / 2)
 Y = Int((.ScaleHeight - .TextHeight(Text1)) / 2 - BorderSize)

 'If Place was selected CboAlign.ListIndex has now -1 so we dont change it
 'even if font has changed
 If LMBAlign(1).IsButtonDown Then
  TextX = X
  TextY = Y
  ElseIf LMBAlign(0).IsButtonDown Then
  TextX = 0 'y is still center
  TextY = Y
  ElseIf LMBAlign(2).IsButtonDown Then
  TextX = .ScaleWidth - .TextWidth(Text1) 'y = still center
  TextY = Y
 Else
  Exit Sub
 End If
 'now if not left then move to the left (Bordersize)
 If TextX <> 0 Then
  TextX = TextX - BorderSize
 End If
 End With
End Sub

'Found this somewhere in the Net
'Needed for Transparent Blit
'Windows 98 and higer has a great API for such things
'But i wanted this code to work with 95 too
Private Sub TransparentBlt(OutDstDC, DstDC, SrcDC, SrcRect As RECT, DstX, DstY, TransColor As Long)
 'DstDC=Device context into which image must be drawn transparently
 'OutDstDC=Device context into image is actually drawn, even though it is made transparent in terms of DstDC
 'Src=Device context of source to be made transparent in color TransColor
 'SrcRect=rectangular region within SrcDC to be made transparent in terms of DstDC, and drawn to OutDstDC
 'DstX, DstY =coordinates in OutDstDC (and DstDC) where tranparent bitmap must go

 Rem In most cases, OutDstDC and DstDC will be the same

 Dim nRet As Long, W As Integer, h As Integer
 Dim MonoMaskDC As Long, hMonoMask As Long
 Dim MonoInvDC As Long, hMonoInv As Long
 Dim ResultDstDC As Long, hResultDst As Long
 Dim ResultSrcDC As Long, hResultSrc As Long
 Dim hPrevMask As Long, hPrevInv As Long, hPrevSrc As Long, hPrevDst As Long
 W = SrcRect.Right - SrcRect.Left + 1
 h = SrcRect.Bottom - SrcRect.Top + 1

 'create monochrome mask and inverse masks
 MonoMaskDC = CreateCompatibleDC(DstDC)
 MonoInvDC = CreateCompatibleDC(DstDC)
 hMonoMask = CreateBitmap(W, h, 1, 1, ByVal 0&)
 hMonoInv = CreateBitmap(W, h, 1, 1, ByVal 0&)
 hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
 hPrevInv = SelectObject(MonoInvDC, hMonoInv)

 'create keeper DCs and bitmaps
 ResultDstDC = CreateCompatibleDC(DstDC)
 ResultSrcDC = CreateCompatibleDC(DstDC)
 hResultDst = CreateCompatibleBitmap(DstDC, W, h)
 hResultSrc = CreateCompatibleBitmap(DstDC, W, h)
 hPrevDst = SelectObject(ResultDstDC, hResultDst)
 hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)

 'copy src to monochrome mask
 Dim OldBC As Long
 OldBC = SetBkColor(SrcDC, TransColor)
 nRet = BitBlt(MonoMaskDC, 0, 0, W, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
 TransColor = SetBkColor(SrcDC, OldBC)

 'create inverse of mask
 nRet = BitBlt(MonoInvDC, 0, 0, W, h, MonoMaskDC, 0, 0, vbNotSrcCopy)

 'get background
 nRet = BitBlt(ResultDstDC, 0, 0, W, h, DstDC, DstX, DstY, vbSrcCopy)
 'AND with Monochrome mask
 nRet = BitBlt(ResultDstDC, 0, 0, W, h, MonoMaskDC, 0, 0, vbSrcAnd)
 'get overlapper
 nRet = BitBlt(ResultSrcDC, 0, 0, W, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
 'AND with inverse monochrome mask
 nRet = BitBlt(ResultSrcDC, 0, 0, W, h, MonoInvDC, 0, 0, vbSrcAnd)
 'XOR these two
 nRet = BitBlt(ResultDstDC, 0, 0, W, h, ResultSrcDC, 0, 0, vbSrcInvert)

 'output results
 nRet = BitBlt(OutDstDC, DstX, DstY, W, h, ResultDstDC, 0, 0, vbSrcCopy)

 'clean up
 hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
 DeleteObject hMonoMask
 hMonoInv = SelectObject(MonoInvDC, hPrevInv)
 DeleteObject hMonoInv
 hResultDst = SelectObject(ResultDstDC, hPrevDst)
 DeleteObject hResultDst
 hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
 DeleteObject hResultSrc
 DeleteDC MonoMaskDC
 DeleteDC MonoInvDC
 DeleteDC ResultDstDC
 DeleteDC ResultSrcDC
End Sub

'Create Gradient
'The same like in Transparent Blit Nice API in 98 or higher
Private Sub DrawGradient(Idx As Integer)
 Dim RGBc(1) As RGBcolor
 Dim I As Integer
 Dim X As Single, Y As Single, z As Single, h As Long, j As Long, k As Long
 With RGBc(0) '13 times RGBc(0) sow we take this for with Statement
 'Get the RGB Codes for the choosen Colors
 RGBc(1).R = CmdGrCol(Idx * 2).BackColor And 255
 RGBc(1).G = (CmdGrCol(Idx * 2).BackColor And 65280) \ 256
 RGBc(1).B = (CmdGrCol(Idx * 2).BackColor And 16711680) \ 65535
 .R = CmdGrCol(Idx * 2 + 1).BackColor And 255
 .G = (CmdGrCol(Idx * 2 + 1).BackColor And 65280) \ 256
 .B = (CmdGrCol(Idx * 2 + 1).BackColor And 16711680) \ 65535
 'Fill the background so we dont get white borders
 'around round dradients
 Pic(2).BackColor = CmdGrCol(Idx * 2 + 1).BackColor
 Pic(2).Cls
 'If the font is bigger than our pic we will raise an error
On Error Resume Next
'Circular or Left to Right
If CboGrStyle(Idx).ListIndex < 2 Then
 h = Pic(0).TextWidth(Text1)
 j = (Pic(0).Width - h) / 2
Else
 'Up to Down
 h = Pic(0).TextHeight(Text1)
 j = (Pic(0).Height - h) / 2
End If
If Idx = 2 Then
 If CboGrStyle(Idx).ListIndex = 2 Then
  h = Pic(0).Height
 Else
  If Pic(0).Width > Pic(0).Height Then
   h = Pic(0).Width
  Else
   h = Pic(0).Height
  End If
 End If
 j = 0
End If
'Now we calculate the difference
X = (RGBc(1).R - .R) / h
Y = (RGBc(1).G - .G) / h
z = (RGBc(1).B - .B) / h
k = 0
'Draw the gradient
Select Case CboGrStyle(Idx).ListIndex
Case 0
 k = Int(h / 2)
 X = X * 2
 Y = Y * 2
 z = z * 2
 For I = j + h / 2 To j Step -1
  Pic(2).FillColor = RGB(RGBc(1).R - k * X, RGBc(1).G - k * Y, RGBc(1).B - k * z)
  Pic(2).Circle (Pic(2).Width / 2, Pic(2).Height / 2), k, Pic(2).FillColor
  k = k - 1
 Next I
Case 1
 For I = j To j + h
  Pic(2).Line (I, 0)-(I, Pic(2).Height), RGB(.R + k * X, .G + k * Y, .B + k * z)
  k = k + 1
 Next I
Case 2
 For I = j To j + h
  Pic(2).Line (0, I)-(Pic(2).Width, I), RGB(.R + k * X, .G + k * Y, .B + k * z)
  k = k + 1
 Next I
End Select
End With
End Sub

'Need this for VB5
'Think VB6 has split :(
'Cuts all bevore the last ??
Public Function CutBefore(ByVal StringToCut As String, ByVal CutString As String)
 Dim Cutlenght As Integer
 Cutlenght = Len(StringToCut) - Len(CutString) + 1
 Do Until Mid$(StringToCut, Cutlenght, Len(CutString)) = CutString
  Cutlenght = Cutlenght - 1
 Loop
 CutBefore = Right$(StringToCut, Len(StringToCut) - Cutlenght)
End Function


'Move over Pic(1)
Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Show the Backlight Preview
 If ChkLight.Value = 1 Then
  DrawLight X, Pic(0).Height - Y, True
  Exit Sub
 End If

 With Pic(1)
 'Now drawing function so quit
 If .MousePointer = 0 Or ShpSun.Visible = True Then Exit Sub

 .AutoRedraw = False

 'We are in drawing mode so show the Magnify
 If FrmMainTab(3).Visible = True Then
  'Magnify 10 times
  StretchBlt PicMagnify.hdc, 0, 0, 200, 200, .hdc, X - 10, Y - 10, 20, 20, vbSrcCopy
  'Show the cursorpos in the Mangnify
  PicMagnify.Line (100, 100)-(110, 110), &HFFFFFF, B
 End If

 Select Case Funkt
 Case 1 'Bend horizontal
  'position hasn´t changed so exit
  If PaintY = X Then Exit Sub
  'remember last position
  PaintY = X
  'clear
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  'bend
  BendX X
  Exit Sub
 Case 2 'Bend vertical
  If PaintY = Y Then Exit Sub
  PaintY = Y
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  'BendX y,true sends Cursorpos and True for Vertical
  BendX Y, True
  Exit Sub
 Case 3 'Set Textpos
  Dim rc As RECT
  'Set a rect as big as our text is

  rc.Left = Int((Pic(0).ScaleWidth - Pic(0).TextWidth(Text1)) / 2) - BorderSize
  rc.Right = rc.Left + Pic(0).TextWidth(Text1) + BorderSize
  rc.Top = Int((Pic(0).ScaleHeight - Pic(0).TextHeight(Text1)) / 2) - BorderSize
  rc.Bottom = rc.Top + Pic(0).TextHeight(Text1) + BorderSize
  .Cls
  'now blit text to actual x,y Pos
  TransparentBlt .hdc, .hdc, Pic(3).hdc, rc, X - Pic(0).TextWidth(Text1) / 2, Y - Pic(0).TextHeight(Text1) / 2, &HFFFFFF
 End Select

 'Paint functions
 If PaintX <> -1 Then
  .DrawMode = 7
  If PaintMode = 2 Then 'Line
   BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
   Pic(1).Line (PaintX, PaintY)-(X, Y), &HFFFFFF
   ElseIf PaintMode = 3 Then 'Box
   BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
   Pic(1).Line (PaintX, PaintY)-(X, Y), &HFFFFFF, BF
  End If
  .DrawMode = 13
 End If

 'If point is choosen and button is down set the next point
 If Button <> 0 And PaintMode = 1 Then
  Pic_MouseDown 1, Button, Shift, X, Y
 End If

 End With
End Sub

'Draw
Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 'Draw the Backlight
 If ChkLight.Value = 1 Then
  'Turn Prview off
  DrawLight X, Pic(0).Height - Y, False
  ChkLight.Value = 0
  Exit Sub
 End If

 'Right button Click (Quit)
 With Pic(1)
 If Button = 2 Then
  'draw the
  If Funkt < 4 Then
   'Set back to start (no first poit choosen)
   PaintX = -1
   BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
   'Set Textpos Clear Combo set pic to normal
   If Funkt = 3 Then
    .MousePointer = 0
    .AutoRedraw = True
   End If
   If Funkt < 3 Then
    DrawPic
   End If
  End If
  Funkt = 0
  Exit Sub
 End If

 'We changed the Pic
 ChangedPic = True

 'Draw (Point,Line...)
 Select Case PaintMode
 Case 1 'Point
  .AutoRedraw = True
  Pic(1).PSet (X, Y), PaintColor
  .AutoRedraw = False
 Case 2 'Line
  'If no first point is choosen then remember x,y
  If PaintX = -1 Then
   PaintX = X
   PaintY = Y
  Else
   'now set line from old x,y to actual x,y
   .AutoRedraw = True
   Pic(1).Line (PaintX, PaintY)-(X, Y), PaintColor
   'no first choosen
   PaintX = -1
   .AutoRedraw = False
  End If
 Case 3 'Box (same as line)
  If PaintX = -1 Then
   PaintX = X
   PaintY = Y
  Else
   .AutoRedraw = True
   Pic(1).Line (PaintX, PaintY)-(X, Y), PaintColor, BF
   PaintX = -1
   .AutoRedraw = False
  End If
 Case 4
  PaintColor = .Point(X, Y)
  CmdDrawColor.BackColor = PaintColor
 End Select

 Select Case Funkt
 Case 1 To 2 'Bend
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  .AutoRedraw = True
  If Funkt = 1 Then
   BendX (X)
  Else
   BendX Y, True
  End If
  .AutoRedraw = False
  Funkt = 0
  .Refresh
 Case 3 'Place Text
  TextX = Int(X - Pic(0).TextWidth(Text1) / 2)
  TextY = Int(Y - Pic(0).TextHeight(Text1) / 2)
  .AutoRedraw = True
  DrawPic
  .MousePointer = 0
  ChangedPic = False
  Funkt = 0
 End Select
 End With
 Pic_MouseMove Index, 0, Shift, X, Y
End Sub

'Bend Font Horizontal / Vertical
Private Sub BendX(ByVal X As Long, Optional Vertical As Boolean)
 Dim I As Single, F As Long, G As Long, h As Long
 Dim j As Long, k As Long, l As Long, hfg As Long
 Dim rc As RECT
 Dim Tmp As Byte

 'Check if Invert is on
 'If yes and preview is done then
 'Draw new pic to pic(0) not 1
 If ChkInvert And Pic(1).AutoRedraw = True Then
  Tmp = 0
  BitBlt Pic(0).hdc, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleHeight, 0, 0, 0, WHITENESS
 Else
  Tmp = 1
 End If

 'Get the Fontsize
 With rc
 .Left = Int((Pic(0).ScaleWidth - Pic(0).TextWidth(Text1)) / 2 - BorderSize - 2)
 .Right = .Left + Pic(0).TextWidth(Text1) + 2 * TxtBordersize + 2
 .Top = Int((Pic(0).ScaleHeight - Pic(0).TextHeight(Text1)) / 2 - BorderSize)
 .Bottom = .Top + Pic(0).TextHeight(Text1) + 2 * BorderSize
 F = .Left
 G = .Top
 h = .Right
 j = .Bottom

 'Calculate a Elipse arount the text
 'Size = text Hight & X or Width & Y
 If Vertical Then
  X = X + (.Top - TextY) - (.Bottom - .Top) / 2 - 2
  For I = 181 To 361 Step 180 / (Abs(h - F) * 2)
   k = Cos(I * PI / 180) * (Abs(h - F) / 2)
   l = Sin(I * PI / 180) * (G - X)
   .Left = F + (h - F) / 2 + k
   .Right = .Left + 1
   hfg = Int((Pic(0).ScaleWidth - Pic(0).TextWidth(Text1)) / 2 - TxtBordersize)
   TransparentBlt Pic(Tmp).hdc, Pic(Tmp).hdc, Pic(3).hdc, rc, TextX + .Left - hfg, TextY + (G + l - .Top), &HFFFFFF
  Next
 Else
  X = X + (.Left - TextX) - (.Right - .Left) / 2 - 2
  For I = 91 To 271 Step 180 / (Abs(j - G) * 2)
   k = Cos(I * PI / 180) * (F - X)
   l = Sin(I * PI / 180) * Abs(j - G) / 2
   .Top = Int(l + j - Abs(j - G) / 2)
   .Bottom = .Top + 1
   hfg = Int(((Pic(0).ScaleHeight - Pic(0).TextHeight(Text1)) / 2 - TxtBordersize))
   TransparentBlt Pic(Tmp).hdc, Pic(Tmp).hdc, Pic(3).hdc, rc, TextX + (k + F) - .Left, TextY + (.Top - hfg), &HFFFFFF
  Next
 End If
 End With

 'Invert = on and Prview is done
 'no Invert Pic(0) over pic(1)
 If Tmp = 0 Then
  BltInvert Pic(0), 0, 0, Pic(0).Width, Pic(0).Height, Pic(1), 0, 0, &HFFFFFF
 End If
End Sub

'Scrollbars
Private Sub ChangePicSize()
 Dim SizeX As Integer
 Dim SizeY As Integer
 Dim MaxX As Integer
 Dim MaxY As Integer
 Dim I As Byte
 Dim X  As Long, Y As Long
 'Resize all Pics we need to the same size
 For I = 0 To 3
  If Pic(I).Width <> Val(TxtSizeX) Or Pic(I).Height <> Val(TxtSizeY) Then
   Pic(I).Width = Val(TxtSizeX)
   Pic(I).Height = Val(TxtSizeY)
   Pic(I).Cls
  End If
 Next I
 'Resize our std Rect
 With Re
 .Left = 0
 .Top = 0
 .Right = Pic(1).ScaleWidth
 .Bottom = Pic(1).ScaleHeight
 End With
 'Show or hide Scrollbars (if pic is bigger than form)
 With Pic(1)
ShowScrollBars:
 SizeX = Val(TxtSizeX)
 SizeY = Val(TxtSizeY)
 MaxX = Me.ScaleWidth
 MaxY = Me.ScaleHeight - FrmBlock.Top - FrmBlock.Height

 If SizeX > MaxX Then
  HScrPic.Visible = True
  SizeY = SizeY + 17
 Else
  HScrPic.Visible = False
 End If

 If SizeY > MaxY Then
  VScrPic.Visible = True
  SizeX = SizeX + 17
 Else
  VScrPic.Visible = False
 End If

 If SizeX > MaxX And HScrPic.Visible = False Then
  HScrPic.Visible = True
  SizeY = SizeY + 17
 End If

 HScrPic.Top = Me.ScaleHeight - 17
 VScrPic.Left = Me.ScaleWidth - 17
 HScrPic.Max = SizeX - MaxX
 VScrPic.Max = SizeY - MaxY
 VScrPic.Height = MaxY


 If VScrPic.Visible = True Then
  HScrPic.Width = MaxX - 17
 Else
  HScrPic.Width = MaxX
 End If


 End With

 'Set new Maxlight for Backlight
 'Maximal distance for the light is the the smallest picturesize
 If Pic(1).Height < Pic(1).Width Then
  HScrDist.Max = Pic(1).Width * 2
  HScrDist.Value = HScrDist.Max
 Else
  HScrDist.Max = Pic(1).Height * 2
  HScrDist.Value = HScrDist.Max
 End If

 SetMovePic
End Sub
Private Sub HScrPic_Change()
 Pic(1).Left = -HScrPic.Value
End Sub
Private Sub VScrPic_Change()
 Pic(1).Top = FrmBlock.Height - VScrPic.Value
End Sub
Private Sub DrawShadow(Optional Real As Boolean)
 Dim X As Double
 Dim Y As Double
 Dim I As Integer
 Dim F As Double
 Dim G As Double, h As Double
 Dim Buf()       As RGBQUAD      'This Array will hold the RGB Colors from Source
 Dim Tmp As Boolean

 'If we draw a real shadow (No Preview)
 If Real Then
  With Pic(0)
  'copy the TextPicture to pic(2)
  X = Int((.ScaleWidth - .TextWidth(Text1)) / 2)
  Y = Int((.ScaleHeight - .TextHeight(Text1)) / 2)
  'Get the coordinates for TransBlit
  Dim rc As RECT
  rc.Left = X - BorderSize
  rc.Right = rc.Left + .TextWidth(Text1) + BorderSize
  rc.Top = Y - BorderSize
  rc.Bottom = rc.Top + .TextHeight(Text1) + BorderSize

  'Check if Anti Alias is selectet
  If ChkAlias.Value = 1 Then
   Tmp = True
  End If
  BitBlt Pic(2).hdc, 0, 0, Pic(2).ScaleWidth, Pic(2).ScaleHeight, 0, 0, 0, WHITENESS
  TranBltAlias Pic(3), rc.Left, rc.Top, .TextWidth(Text1) + BorderSize * 2, .TextHeight(Text1) + BorderSize * 2, Pic(2), TextX, TextY, &HFFFFFF, False, 100

  'Get the Color Pic of our Font
  Pic2Array Pic(2), Buf()

  I = .Height
  'Now make BW
  For X = 0 To Pic(3).Width - 1
   For Y = 0 To Pic(3).Height - 1
    If Buf(X, Y).rgbBlue <> 255 Or Buf(X, Y).rgbGreen <> 255 Or Buf(X, Y).rgbRed <> 255 Then
     'We need this later as Starting Point
     If I > Y Then I = Y

     'Make all nontransparent Pixels Black
     Buf(X, Y).rgbBlue = 0
     Buf(X, Y).rgbGreen = 0
     Buf(X, Y).rgbRed = 0
    End If
   Next Y
  Next X

  'Write new BW pic to Pic(2)
  BitBlt Pic(2).hdc, 0, 0, Pic(2).ScaleWidth, Pic(2).ScaleHeight, 0, 0, 0, WHITENESS
  Array2Pic Pic(2), Buf

  'Now that we have a BW Pic for Shadow we create the New form
  'Get the Form of the new Shadow
  X = 100 - ShpSun.Left
  F = (200 - ShpSun.Top) / 2
  If F = 0 Then F = 0.1
  F = (Pic(2).Height - I) / F
  X = X / ((Pic(2).Height - I))
  G = 0
  Y = Pic(2).Height - I
  h = 0
  F = -(F / 6)
  'Draw it
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  Do Until Y <= 0
   BitBlt .hdc, h, (.Height - I) - G, .Width, 1, Pic(2).hdc, 0, Y, vbSrcCopy
   G = G + 1
   Y = Y + F
   h = h + X
  Loop

  'Draw the Background, the Shadow, the Font
  Pic(1).AutoRedraw = True
  Pic(1).Cls

  DrawBackpic
  If ChkAlpha.Value = 1 Then
   X = Val(TxtAlpha)
  Else
   X = 100
  End If
  TranBltAlias Pic(0), 0, 0, .Width, .Height, Pic(1), 0, 0, &HFFFFFF, Tmp, X / 2
  'TranBltAlias Pic(0), rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, Pic(1), rc.Left - g, rc.Top - h, &HFFFFFF, tmp, x / 2
  'Check Alpha Blending
  If ChkAlpha.Value = 1 Then
   X = Val(TxtAlpha)
  Else
   X = 100
  End If
  'New Check if Anti Alias is selectet
  If ChkAlias.Value = 1 Or ChkAlpha.Value = 1 Then
   TranBltAlias Pic(3), rc.Left, rc.Top, .TextWidth(Text1) + BorderSize * 2, .TextHeight(Text1) + BorderSize * 2, Pic(1), TextX, TextY, &HFFFFFF, Tmp, X
  Else
   TransparentBlt Pic(1).hdc, Pic(1).hdc, Pic(3).hdc, rc, TextX, TextY, &HFFFFFF
  End If
  Pic(1).Refresh
  Pic(1).AutoRedraw = True
  End With

  'Only Preview Mode
 Else
  With Pic(2)
  'Draw the word Shadow on pic 2
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  .ForeColor = &HE0E0E0
  .CurrentX = 0
  .CurrentY = 0
  Pic(2).Print "Shadow"

  'Now calculate the form of the Shadow
  X = 100 - ShpSun.Left
  F = (200 - ShpSun.Top) / 2
  If F = 0 Then F = 0.1
  F = 13 / F
  X = X / 26
  G = 0
  Y = 13
  h = 0
  F = -F
  'Clear Magnify Picture and copy Shadow + Word to it
  'PicMagnify.Cls
  BitBlt PicMagnify.hdc, 0, 0, PicMagnify.ScaleWidth, PicMagnify.ScaleHeight, 0, 0, 0, WHITENESS
  Do Until Y <= 0
   BitBlt PicMagnify.hdc, 65 + h, 200 - G, 100, 1, .hdc, 0, 5 + Y, vbSrcCopy
   G = G + 1
   Y = Y + F
   h = h + X
  Loop
  PicMagnify.CurrentX = 65
  PicMagnify.CurrentY = 181
  PicMagnify.Print "Shadow"
  End With
 End If
End Sub
Private Sub DrawHover(Optional Real As Boolean)
 Dim X As Double
 Dim Y As Double
 Dim I As Integer
 Dim F As Double
 Dim G As Double, h As Double
 Dim Buf()       As RGBQUAD      'This Array will hold the RGB Colors from Source
 Dim Tmp As Boolean

 'If we draw a real shadow (No Preview)
 If Real Then
  With Pic(0)
  'copy the TextPicture to pic(2)
  X = Int((.ScaleWidth - .TextWidth(Text1)) / 2)
  Y = Int((.ScaleHeight - .TextHeight(Text1)) / 2)
  'Get the coordinates for TransBlit
  Dim rc As RECT
  rc.Left = X - BorderSize
  rc.Right = rc.Left + .TextWidth(Text1) + BorderSize
  rc.Top = Y - BorderSize
  rc.Bottom = rc.Top + .TextHeight(Text1) + BorderSize

  'Check if Anti Alias is selectet
  If ChkAlias.Value = 1 Then
   Tmp = True
  End If
  BitBlt Pic(2).hdc, 0, 0, Pic(2).ScaleWidth, Pic(2).ScaleHeight, 0, 0, 0, WHITENESS
  TranBltAlias Pic(3), rc.Left, rc.Top, .TextWidth(Text1) + BorderSize * 2, .TextHeight(Text1) + BorderSize * 2, Pic(2), TextX, TextY, &HFFFFFF, False, 100

  'Get the Color Pic of our Font
  Pic2Array Pic(2), Buf()

  I = .Height
  'Now make BW
  For X = 0 To Pic(3).Width - 1
   For Y = 0 To Pic(3).Height - 1
    If Buf(X, Y).rgbBlue <> 255 Or Buf(X, Y).rgbGreen <> 255 Or Buf(X, Y).rgbRed <> 255 Then
     'We need this later as Starting Point
     If I > Y Then I = Y

     'Make all nontransparent Pixels Black
     Buf(X, Y).rgbBlue = 0
     Buf(X, Y).rgbGreen = 0
     Buf(X, Y).rgbRed = 0
    End If
   Next Y
  Next X

  'Write new BW pic to Pic(2)
  BitBlt Pic(2).hdc, 0, 0, Pic(2).ScaleWidth, Pic(2).ScaleHeight, 0, 0, 0, WHITENESS
  Array2Pic Pic(2), Buf

  G = (ShpSun.Left / 100 - 1) * 15
  h = (ShpSun.Top / 100 - 1) * 15

  'Draw the Background, the Shadow, the Font
  Pic(1).AutoRedraw = True
  Pic(1).Cls

  DrawBackpic
  If ChkAlpha.Value = 1 Then
   X = Val(TxtAlpha)
  Else
   X = 100
  End If
  TranBltAlias Pic(2), rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, Pic(1), rc.Left - G, rc.Top - h, &HFFFFFF, Tmp, X / 2
  'Check Alpha Blending
  If ChkAlpha.Value = 1 Then
   X = Val(TxtAlpha)
  Else
   X = 100
  End If
  'New Check if Anti Alias is selectet
  If ChkAlias.Value = 1 Or ChkAlpha.Value = 1 Then
   TranBltAlias Pic(3), rc.Left, rc.Top, .TextWidth(Text1) + BorderSize * 2, .TextHeight(Text1) + BorderSize * 2, Pic(1), TextX, TextY, &HFFFFFF, Tmp, X
  Else
   TransparentBlt Pic(1).hdc, Pic(1).hdc, Pic(3).hdc, rc, TextX, TextY, &HFFFFFF
  End If
  Pic(1).Refresh
  Pic(1).AutoRedraw = True
  End With

  'Only Preview Mode
 Else
  With PicMagnify
  BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, WHITENESS
  .ForeColor = &HC0C0C0
  X = (ShpSun.Left / 100 - 1) * 10
  Y = (ShpSun.Top / 100 - 1) * 10
  G = 100 - .TextHeight("Hover") / 2
  h = 100 - .TextWidth("Hover") / 2
  .CurrentX = h - X
  .CurrentY = G - Y
  PicMagnify.Print "Hover"
  .CurrentX = h
  .CurrentY = G
  .ForeColor = 0
  PicMagnify.Print "Hover"
  End With
 End If
End Sub

'Draw/Preview the Backlight
'This sub isnt optimized
'i only transfered it from
'an old testproject to LogoMan
Private Sub DrawLight(StartX As Single, StartY As Single, Preview As Boolean)
 Dim X As Single
 Dim Y As Single
 Dim Xp As Single
 Dim Yp As Single
 Dim Multi As Integer
 Dim Mul As Integer
 Dim LghtR As Single
 Dim LghtG As Single
 Dim LghtB As Single
 Dim Lng As Single
 Dim I As Single
 Dim F As Single
 Dim EndX As Single
 Dim EndY As Single
 Dim h As Long
 Dim Buf() As RGBQUAD
 Dim Buf1() As RGBQUAD
 Dim Stp As Single
 Dim MulOld As Boolean
 Dim GotPxl As Boolean


 Dim Colr As Byte
 Dim ColG As Byte
 Dim ColB As Byte
 Dim Col As Long

 'Refresh the Picture
 DrawPic

 'Draw the Font in BW
 SystemParametersInfo SPI_GETFONTSMOOTHING, 0, h, 0
 SystemParametersInfo SPI_SETFONTSMOOTHING, 0, 0&, SPIF_SENDWININICHANGE
 Lng = Pic(0).ForeColor
 GetTextPos
 'Pic(0).Cls
 BitBlt Pic(0).hdc, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleHeight, 0, 0, 0, WHITENESS
 Pic(0).CurrentX = TextX + BorderSize
 Pic(0).CurrentY = TextY + BorderSize
 Pic(0).ForeColor = &H0
 Pic(0).Print Text1
 Pic(0).ForeColor = Lng
 SystemParametersInfo SPI_SETFONTSMOOTHING, 0, h, SPIF_SENDWININICHANGE
 h = 0

 If Preview = True Then
  Stp = 1
  Pic2Array Pic(1), Buf1()  'To Paint on
  Colr = CmdLightCol.BackColor And 255
  ColG = (CmdLightCol.BackColor And 65280) \ 256
  ColB = (CmdLightCol.BackColor And 16711680) \ 65535
 Else
  Stp = 0.1
  'Create a complete black Picture
  ReDim Buf1(0 To Pic(0).ScaleWidth - 1, 0 To Pic(0).ScaleHeight - 1)
  Colr = 255
  ColB = 255
  ColG = 255
 End If

 'Get the Picture into our array
 Pic2Array Pic(0), Buf() 'To scan the Picture

 On Error GoTo Done

 For F = 0 To 360 Step Stp
  EndX = SinCos(h).X * HScrDist.Value + StartX
  EndY = SinCos(h).Y * HScrDist.Value + StartY
  h = h + 10 * Stp

  'Calculate the Distance between Start & End
  Lng = Distance(StartX, StartY, EndX, EndY)

  '  If Lng = 0 Then Exit For

  'Calculate the difference between every new point
  Xp = (EndX - StartX) / Lng
  Yp = (EndY - StartY) / Lng

  'Set Starting Point
  X = StartX
  Y = StartY

  'Now draw the line
  For I = 0 To Lng

   'Check if we are off the Picture
   If X < 0 Then Exit For
   If X > Pic(0).Width - 1 Then Exit For
   If Y < 0 Then Exit For
   If Y > Pic(0).Height - 1 Then Exit For

   'If we have a Fontpixel then
   If Buf(X, Y).rgbBlue = 0 Then
    'Add x new Points to light
    Multi = Multi + HScrStrnght.Value
    'If we want real light then shine over the font
    'else start new after the font
    If ChkReal.Value = 1 Then
     GotPxl = True
     If Mul = 0 Then Mul = 100: MulOld = False
    Else
     Mul = 0
     'GotPxl = False
    End If
   Else
    'Real mode ? If yes start new after the font
    If GotPxl = True Then GotPxl = False: Mul = 0
    'If we have new Points
    If Multi > 0 Then
     'Take them for drawing
     Mul = Multi
     MulOld = False
     'set new to 0 so we dont change until we reach new fontparts
     Multi = 0
    End If
   End If

   'Calculate the Multiplicator for distance
   If I + Mul > Lng Then Mul = Lng - I
   If Mul > 0 And MulOld = False Then
    LghtR = (Colr - I) / Mul
    If LghtR < 0 Then LghtR = 0
    LghtG = (ColG - I) / Mul
    If LghtG < 0 Then LghtG = 0
    LghtB = (ColB - I) / Mul
    If LghtB < 0 Then LghtB = 0
    MulOld = True
   End If


   'If we need to draw new light
   If Mul > 0 Then
    'Draw it
    Buf1(X, Y).rgbBlue = CLng(LghtB * Mul)
    'If Preview Or Check1.Value = 0 Then
    Buf1(X, Y).rgbRed = CLng(LghtR * Mul)
    Buf1(X, Y).rgbGreen = CLng(LghtG * Mul)
    'End If
    'Reduce distance
    Mul = Mul - 1
   End If

   'Move to next point
   X = X + Xp
   Y = Y + Yp

  Next I
Done:
  Multi = 0
  Mul = 0
  MulOld = False
 Next F

 If Preview = True Then
  Array2Pic Pic(1), Buf1()
 Else
  'Get the Picture to set the light on it
  Pic2Array Pic(1), Buf()

  'Get the color we use for the light
  Colr = CmdLightCol.BackColor And 255
  ColG = (CmdLightCol.BackColor And 65280) \ 256
  ColB = (CmdLightCol.BackColor And 16711680) \ 65535

  'Scann our new picture an calculate the light if needed
  For I = 0 To Pic(0).ScaleWidth - 1
   For F = 0 To Pic(0).ScaleHeight - 1
    'No black Point so do something
    If Buf1(I, F).rgbBlue <> 0 Then
     Col = CLng(Buf(I, F).rgbRed + CLng(Buf1(I, F).rgbBlue) / 255 * Colr)
     If Col > 255 Then Col = 255
     Buf(I, F).rgbRed = Col
     Col = CLng(Buf(I, F).rgbGreen + CLng(Buf1(I, F).rgbBlue) / 255 * ColG)
     If Col > 255 Then Col = 255
     Buf(I, F).rgbGreen = Col
     Col = CLng(Buf(I, F).rgbBlue + CLng(Buf1(I, F).rgbBlue) / 255 * ColB)
     If Col > 255 Then Col = 255
     Buf(I, F).rgbBlue = Col
    End If
   Next F
  Next I
  Array2Pic Pic(1), Buf()
 End If

 'Show our new Picture
 Pic(1).Refresh
End Sub

'Needed for Baclight distance calculations
Public Function Distance(StartX As Single, StartY As Single, EndX As Single, EndY As Single) As Single
 Distance = Sqr(((EndX - StartX) ^ 2) + ((EndY - StartY) ^ 2))
End Function

'Transparent Blit with Anti Alias and Alpha Blend
'I´m still Optimizing this Routine
Private Function TranBltAlias(Source As PictureBox, ByVal SourceX As Integer, ByVal SourceY As Integer, ByVal SizeX As Integer, ByVal SizeY As Integer, Destination As PictureBox, ByVal DestX As Integer, ByVal DestY As Integer, TransColor As Long, ByVal AliasOn As Boolean, Optional ByVal AlphaPercent As Double)
On Error Resume Next
'We use Arrays to Manipulate the Picture
Dim bufS()       As RGBQUAD      'This Array will hold the RGB Colors from Source
Dim bufD()       As RGBQUAD      'This Array will hold the RGB Colors from Dest
Dim Bw()         As Boolean      'Create a BW Pic (Transparent = False / Color = True)
Dim X            As Long         'Needed to move thru the Array
Dim Y            As Long         'Needed to move thru the Array
Dim Col          As Long         'Color Calculations

'Holds the RGB Colors for Transparent Color
Dim Colr         As Byte
Dim ColG         As Byte
Dim ColB         As Byte
Dim Trans As Double


'Correct to big logos
If SourceY < 0 Then SourceY = 0
If SizeY > Source.ScaleHeight Then SizeY = Source.ScaleHeight - 1

'Calculate the Real Position
DestY = Int((Destination.Height - SizeY) / 2) * 2 - DestY

'Calculate Alpha if selectect
AlphaPercent = AlphaPercent / 100

'Get RGB Values for Transparent Color
Colr = TransColor And 255
ColG = (TransColor And 65280) \ 256
ColB = (TransColor And 16711680) \ 65535

'Now get the 2 Pictures
Pic2Array Source, bufS()
Pic2Array Destination, bufD()

'We need the BW Pic only if Alias is on
ReDim Bw(0 To Source.Width - 1, 0 To Source.Height - 1)
'Create the BW Picture
For X = 0 To SizeX '- 1
 For Y = 0 To SizeY
  If bufS(SourceX + X, SourceY + Y).rgbBlue <> ColB Or bufS(SourceX + X, SourceY + Y).rgbGreen <> ColG Or bufS(SourceX + X, SourceY + Y).rgbRed <> Colr Then
   Bw(X, Y) = True
  End If
 Next Y
Next X

For X = 0 To SizeX '- 1
 For Y = 0 To SizeY
  'Check if the Pixel is Transparent
  If Bw(X, Y) = True Then
   'AntiAlias
   If AliasOn Then
    'If not then check the Pixels Left, Right, Up, Down for tranparency
    If Bw(X - 1, Y) = True Then
     Trans = 1
    End If
    If Bw(X + 1, Y) = True Then
     Trans = Trans + 1
    End If
    If Bw(X, Y - 1) = True Then
     Trans = Trans + 1
    End If
    If Bw(X, Y + 1) = True Then
     Trans = Trans + 1
    End If
    'Now get the transparecy
    Trans = Trans / 4 * AlphaPercent
   Else
    Trans = AlphaPercent
   End If

   'Normal copy No transparence
   If Trans = 1 Then
    bufD(DestX + X, DestY + Y).rgbBlue = CLng(bufS(SourceX + X, SourceY + Y).rgbBlue)
    bufD(DestX + X, DestY + Y).rgbGreen = CLng(bufS(SourceX + X, SourceY + Y).rgbGreen)
    bufD(DestX + X, DestY + Y).rgbRed = CLng(bufS(SourceX + X, SourceY + Y).rgbRed)
   Else
    'Calculate the color we need to set
    Col = CLng(bufD(DestX + X, DestY + Y).rgbBlue * (1 - Trans) + Trans * CLng(bufS(SourceX + X, SourceY + Y).rgbBlue))
    If Col > 255 Then Col = 255
    bufD(DestX + X, DestY + Y).rgbBlue = Col
    Col = CLng(bufD(DestX + X, DestY + Y).rgbGreen * (1 - Trans) + Trans * CLng(bufS(SourceX + X, SourceY + Y).rgbGreen))
    If Col > 255 Then Col = 255
    bufD(DestX + X, DestY + Y).rgbGreen = Col
    Col = CLng(bufD(DestX + X, DestY + Y).rgbRed * (1 - Trans) + Trans * CLng(bufS(SourceX + X, SourceY + Y).rgbRed))
    If Col > 255 Then Col = 255
    bufD(DestX + X, DestY + Y).rgbRed = Col
    Trans = 0
   End If
  End If
 Next Y
Next X

'Now show the result
Array2Pic Destination, bufD
End Function
Private Function BltInvert(Source As PictureBox, ByVal SourceX As Integer, ByVal SourceY As Integer, ByVal SizeX As Integer, ByVal SizeY As Integer, Destination As PictureBox, ByVal DestX As Integer, ByVal DestY As Integer, TransColor As Long)
On Error Resume Next
'We use Arrays to Manipulate the Picture
Dim bufS()       As RGBQUAD      'This Array will hold the RGB Colors from Source
Dim bufD()       As RGBQUAD      'This Array will hold the RGB Colors from Dest
Dim X            As Long         'Needed to move thru the Array
Dim Y            As Long         'Needed to move thru the Array

'Holds the RGB Colors for Transparent Color
Dim Colr         As Byte
Dim ColG         As Byte
Dim ColB         As Byte

'Calculate the Real Position
DestY = Int((Destination.Height - SizeY) / 2) * 2 - DestY

Pic2Array Source, bufS()
Pic2Array Destination, bufD()

'Get RGB Values for Transparent Color
Colr = TransColor And 255
ColG = (TransColor And 65280) \ 256
ColB = (TransColor And 16711680) \ 65535

For X = 0 To SizeX '- 1
 For Y = 0 To SizeY
  If bufS(SourceX + X, SourceY + Y).rgbBlue <> ColB Or bufS(SourceX + X, SourceY + Y).rgbGreen <> ColG Or bufS(SourceX + X, SourceY + Y).rgbRed <> Colr Then
   bufD(DestX + X, DestY + Y).rgbBlue = 255 - bufD(DestX + X, DestY + Y).rgbBlue
   bufD(DestX + X, DestY + Y).rgbGreen = 255 - bufD(DestX + X, DestY + Y).rgbGreen
   bufD(DestX + X, DestY + Y).rgbRed = 255 - bufD(DestX + X, DestY + Y).rgbRed
  End If
 Next Y
Next X

Array2Pic Destination, bufD()
End Function

'Convert Picture to Array
Private Sub Pic2Array(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 ReDim PicArray(0 To PicBox.ScaleWidth - 1, 0 To PicBox.ScaleHeight - 1)
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 'Now get the Picture
 GetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS
End Sub

'Convert Array to Picture
Private Sub Array2Pic(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 SetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS
End Sub

Private Function GetAsString() As String
 Dim Tmp As String
 Line Input #1, Tmp
 GetAsString = Tmp
End Function
Private Function GetAsLong() As Long
 Dim Tmp As String
 Line Input #1, Tmp
 GetAsLong = Val(Tmp)
End Function
Private Function CutAfter(ByVal ToCut As String, ByVal CutString As String)
 Dim Cutlenght As Long
 Cutlenght = Len(ToCut) - Len(CutString) + 1
 Do Until Mid$(ToCut, Cutlenght, Len(CutString)) = CutString
  Cutlenght = Cutlenght - 1
 Loop
 CutAfter = Left$(ToCut, Cutlenght)
End Function

'Shows the Previewpicture in cmdlg
Public Sub ShowPreview(FName As String)
 Dim I As Long
 Dim Pos As Long
 Dim Tmp As Long
 Dim Buf() As Byte
 Dim FLen As Long
 Dim Vers As Byte
 'On Error GoTo ErrOut

 'is the file we want to load there ?
 If Dir$(FName) <> "" Then
  FLen = FileLen(FName)
  With FrmCmdlg.PicPreview
  'Open file
  .Cls
  Open FName For Binary Access Read As #1
   Get #1, , Vers
   Get #1, , Pos 'Get Preview pic position
   FLen = FLen - Pos - 8
   Seek #1, Pos   'move there
   Get #1, , Tmp  'get picture height and resize prev pic
   .Height = Tmp
   .Top = (FrmCmdlg.ScaleHeight - Tmp) / 2
   Get #1, , Tmp 'width
   .Width = Tmp
   .Left = (FrmCmdlg.ScaleWidth - Tmp) / 2
   ReDim Buf(FLen)   'create the buffer to load pic
   Get #1, , Buf() 'load pic
   'if fileversion >1 then Preview is compressed
   If Vers > 1 Then DeCompressArray Buf
   Array2Pic1d FrmCmdlg.PicPreview, Buf() 'show
   .Refresh
   End With
 End If
ErrOut:
 Close #1
End Sub

'Convert Picture to Bytearray
Private Sub Pic2Array1D(PicBox As PictureBox, ByRef PicArray() As Byte)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 ReDim PicArray((PicBox.ScaleWidth) * (PicBox.ScaleHeight) * 4)
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 'Now get the Picture
 GetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0), Binfo, DIB_RGB_COLORS
End Sub

'Convert Bytearray to Picture
Private Sub Array2Pic1d(PicBox As PictureBox, ByRef PicArray() As Byte)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 SetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0), Binfo, DIB_RGB_COLORS
End Sub

