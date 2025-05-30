VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form dockSettings 
   Appearance      =   0  'Flat
   Caption         =   "SteamyDock Settings"
   ClientHeight    =   9645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "dockSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   8790
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   256
      ToolTipText     =   "This will save your changes and restart the dock."
      Top             =   9030
      Width           =   780
   End
   Begin VB.Timer positionTimer 
      Interval        =   3000
      Left            =   1455
      Top             =   9090
   End
   Begin VB.CheckBox chkToggleDialogs 
      Caption         =   "Display Info.Dialogs"
      Height          =   225
      Left            =   1650
      TabIndex        =   210
      ToolTipText     =   "When checked this toggle will display the information pop-ups and balloon tips "
      Top             =   8745
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   150
      ToolTipText     =   "Click here to open tool's HTML help page in your browser"
      Top             =   9030
      Width           =   1065
   End
   Begin VB.PictureBox picBusy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4635
      Picture         =   "dockSettings.frx":058A
      ScaleHeight     =   795
      ScaleWidth      =   825
      TabIndex        =   149
      ToolTipText     =   "The program is doing something..."
      Top             =   8835
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Timer busyTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   990
      Top             =   8910
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3915
      Top             =   9060
   End
   Begin VB.Timer repaintTimer 
      Interval        =   1000
      Left            =   5760
      Top             =   8745
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit this utility"
      Top             =   9030
      Width           =   900
   End
   Begin VB.CommandButton btnSaveRestart 
      Caption         =   "Save+&Restart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "This will save your changes and restart the dock."
      Top             =   9030
      Width           =   1260
   End
   Begin VB.PictureBox iconBox 
      BackColor       =   &H00FFFFFF&
      Height          =   9285
      Left            =   75
      ScaleHeight     =   9225
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   135
      Width           =   1395
      Begin VB.Frame fmeWallpaper 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   15
         TabIndex        =   236
         Top             =   6660
         Width           =   1605
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Wallpaper"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   237
            Top             =   1005
            Width           =   750
         End
         Begin VB.Image imgIcon 
            Height          =   840
            Index           =   5
            Left            =   210
            Picture         =   "dockSettings.frx":1005
            Top             =   0
            Width           =   840
         End
         Begin VB.Image imgIconPressed 
            Height          =   840
            Index           =   5
            Left            =   225
            Picture         =   "dockSettings.frx":2725
            Top             =   15
            Width           =   840
         End
      End
      Begin VB.Frame fmeAbout 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   30
         TabIndex        =   15
         Top             =   7980
         Width           =   1590
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   345
            TabIndex        =   16
            Top             =   960
            Width           =   570
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   6
            Left            =   195
            Picture         =   "dockSettings.frx":3CF5
            Top             =   -15
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   6
            Left            =   210
            Picture         =   "dockSettings.frx":514C
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.Frame fmeGeneral 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   60
         TabIndex        =   14
         Top             =   -75
         Width           =   1350
         Begin VB.Frame fmeLblGeneral 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   60
            TabIndex        =   18
            Top             =   1050
            Width           =   930
            Begin VB.Label lblText 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "General"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   19
               ToolTipText     =   "General Configuration Options"
               Top             =   30
               Width           =   750
            End
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   0
            Left            =   120
            Picture         =   "dockSettings.frx":651C
            Top             =   165
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   0
            Left            =   135
            Picture         =   "dockSettings.frx":7A3C
            Top             =   180
            Width           =   960
         End
      End
      Begin VB.Frame fmePosition 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   15
         TabIndex        =   13
         Top             =   5130
         Width           =   1605
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   390
            TabIndex        =   235
            Top             =   930
            Width           =   765
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   4
            Left            =   165
            Picture         =   "dockSettings.frx":8E83
            Top             =   0
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   4
            Left            =   180
            Picture         =   "dockSettings.frx":A447
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.Frame fmeStyle 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1170
         Left            =   0
         TabIndex        =   6
         Top             =   3840
         Width           =   1515
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Style"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   450
            TabIndex        =   17
            Top             =   930
            Width           =   765
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   3
            Left            =   150
            Picture         =   "dockSettings.frx":B9B4
            Top             =   0
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   3
            Left            =   150
            Picture         =   "dockSettings.frx":CE64
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.Frame fmeBehaviour 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   30
         TabIndex        =   5
         Top             =   2535
         Width           =   1425
         Begin VB.Frame fmeLblBehaviour 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   15
            TabIndex        =   20
            Top             =   990
            Width           =   1020
            Begin VB.Label lblText 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Behaviour"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   210
               TabIndex        =   21
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   2
            Left            =   150
            Picture         =   "dockSettings.frx":E1AF
            Top             =   0
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   2
            Left            =   165
            Picture         =   "dockSettings.frx":F5E6
            Top             =   15
            Width           =   960
         End
      End
      Begin VB.Frame fmeIcons 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   30
         TabIndex        =   4
         Top             =   1335
         Width           =   1590
         Begin VB.Label lblText 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Icons"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   450
            TabIndex        =   255
            Top             =   915
            Width           =   375
         End
         Begin VB.Image imgIcon 
            Height          =   960
            Index           =   1
            Left            =   150
            Picture         =   "dockSettings.frx":100F0
            Top             =   0
            Width           =   960
         End
         Begin VB.Image imgIconPressed 
            Height          =   960
            Index           =   1
            Left            =   180
            Picture         =   "dockSettings.frx":117CF
            Top             =   15
            Width           =   960
         End
      End
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "De&faults"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Revert ALL settings to the defaults"
      Top             =   9030
      Width           =   1065
   End
   Begin VB.PictureBox picHiddenPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   6525
      ScaleHeight     =   1605
      ScaleWidth      =   1485
      TabIndex        =   75
      ToolTipText     =   "The icon size in the dock"
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Desktop Wallpaper Settings "
      Height          =   8595
      Index           =   5
      Left            =   1665
      TabIndex        =   238
      Top             =   45
      Width           =   6930
      Begin VB.CommandButton btnNextWallpaper 
         Caption         =   ">"
         Height          =   1095
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   3780
         Width           =   300
      End
      Begin VB.CommandButton btnPreviousWallpaper 
         Caption         =   "<"
         Height          =   1095
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   3780
         Width           =   300
      End
      Begin VB.ComboBox cmbWallpaperTimerInterval 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1905
         TabIndex        =   246
         Text            =   "5"
         Top             =   2220
         Width           =   1680
      End
      Begin VB.CheckBox chkAutomaticWallpaperChange 
         Caption         =   "Enable Automatic Wallpaper Change"
         Height          =   300
         Left            =   1905
         TabIndex        =   245
         Top             =   1695
         Width           =   3360
      End
      Begin VB.CommandButton btnApplyWallpaper 
         Caption         =   "&Change  *"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   243
         ToolTipText     =   "This will save your changes and restart the dock."
         Top             =   6405
         Width           =   1335
      End
      Begin VB.ComboBox cmbWallpaper 
         Height          =   315
         Left            =   1905
         TabIndex        =   242
         Text            =   "wallpaper1"
         Top             =   675
         Width           =   4245
      End
      Begin VB.ComboBox cmbWallpaperStyle 
         Height          =   315
         Left            =   1905
         TabIndex        =   239
         Text            =   "Centre"
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Frame fmeWallpaperPreview 
         Height          =   3720
         Left            =   525
         TabIndex        =   250
         Top             =   2580
         Width           =   5910
         Begin VB.Image imgWallpaperPreview 
            BorderStyle     =   1  'Fixed Single
            Height          =   3300
            Left            =   150
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5595
         End
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Auto-Timer:"
         Height          =   300
         Index           =   5
         Left            =   705
         TabIndex        =   249
         Top             =   1725
         Width           =   1275
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "(DaysHhrs/Mins)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3810
         TabIndex        =   248
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Interval :"
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   720
         TabIndex        =   247
         Top             =   2235
         Width           =   1275
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Use buttons or  mouse scrollwheel up + down to preview the available wallpapers."
         Height          =   555
         Index           =   2
         Left            =   720
         TabIndex        =   244
         Top             =   6465
         Width           =   4080
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Wallpaper :"
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   241
         Top             =   705
         Width           =   1275
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Positioning :"
         Height          =   300
         Index           =   0
         Left            =   705
         TabIndex        =   240
         Top             =   1230
         Width           =   1275
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Icon && Dock Behaviour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   2
      Left            =   1665
      TabIndex        =   53
      ToolTipText     =   "Here you can control the behaviour of the animation effects"
      Top             =   45
      Width           =   6930
      Begin VB.ComboBox cmbBehaviourSoundSelection 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":12EA2
         Left            =   2190
         List            =   "dockSettings.frx":12EAF
         TabIndex        =   221
         Text            =   "None"
         Top             =   6150
         Width           =   2620
      End
      Begin VB.CheckBox chkRetainIcons 
         Caption         =   "Retain Original Icons when dragging to the dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2190
         TabIndex        =   215
         Top             =   5610
         Width           =   4455
      End
      Begin VB.CheckBox chkLockIcons 
         Caption         =   "Disable Drag/Drop and Icon Deletion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2190
         TabIndex        =   213
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
         Top             =   5130
         Width           =   4500
      End
      Begin VB.ComboBox cmbHidingKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":12ED3
         Left            =   2190
         List            =   "dockSettings.frx":12EFE
         TabIndex        =   204
         Text            =   "F11"
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4515
         Width           =   2620
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   465
         TabIndex        =   197
         Top             =   3675
         Width           =   6120
         Begin CCRSlider.Slider sliContinuousHide 
            Height          =   315
            Left            =   1575
            TabIndex        =   198
            ToolTipText     =   "Determine how long Steamydock will disappear when told to hide using F11"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   1
            Max             =   120
            Value           =   1
            TickFrequency   =   3
            SelStart        =   1
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "1 min"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   1170
            TabIndex        =   199
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   600
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "Continuous Hide"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   45
            LinkItem        =   "150"
            TabIndex        =   202
            ToolTipText     =   "Determine how long Steamydock will disappear when told to hide for the next few minutes"
            Top             =   -30
            Width           =   1350
         End
         Begin VB.Label lblContinuousHideMsCurrent 
            Caption         =   "(30) mins"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4950
            TabIndex        =   201
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblContinuousHideMsHigh 
            Caption         =   "120m"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4440
            TabIndex        =   200
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   405
         End
      End
      Begin VB.Frame fraAutoHideType 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   375
         TabIndex        =   192
         Top             =   465
         Width           =   5325
         Begin VB.ComboBox cmbAutoHideType 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":12F3F
            Left            =   1770
            List            =   "dockSettings.frx":12F4C
            TabIndex        =   196
            Text            =   "Fade"
            ToolTipText     =   "The type of auto-hide, fade, instant or a slide like Rocketdock"
            Top             =   510
            Width           =   2620
         End
         Begin VB.CheckBox chkAutoHide 
            Caption         =   "On/Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   195
            ToolTipText     =   "You can determine whether the dock will auto-hide or not"
            Top             =   480
            Width           =   2235
         End
         Begin VB.ComboBox cmbIconActivationFX 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":12F66
            Left            =   1770
            List            =   "dockSettings.frx":12F73
            TabIndex        =   193
            Text            =   "Bounce"
            ToolTipText     =   $"dockSettings.frx":12F97
            Top             =   0
            Width           =   2620
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "Icon Attention Effect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   90
            LinkItem        =   "150"
            TabIndex        =   194
            ToolTipText     =   $"dockSettings.frx":1302B
            Top             =   45
            Width           =   1605
         End
      End
      Begin VB.Frame fraAutoHideDuration 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   450
         TabIndex        =   186
         Top             =   1500
         Width           =   6180
         Begin CCRSlider.Slider sliAutoHideDuration 
            Height          =   315
            Left            =   1590
            TabIndex        =   187
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   270
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Min             =   1
            Max             =   5000
            Value           =   1
            TickFrequency   =   100
            SelStart        =   1
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "1ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   1140
            TabIndex        =   191
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblAutoHideDurationMsHigh 
            Caption         =   "5000ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4425
            TabIndex        =   190
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblAutoHideDurationMsCurrent 
            Caption         =   "(250)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5085
            TabIndex        =   189
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "AutoHide Duration"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   45
            LinkItem        =   "150"
            TabIndex        =   188
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   420
         TabIndex        =   180
         Top             =   2175
         Width           =   5805
         Begin CCRSlider.Slider sliBehaviourPopUpDelay 
            Height          =   315
            Left            =   1620
            TabIndex        =   181
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   315
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   1
            Max             =   1000
            Value           =   1
            TickFrequency   =   20
            SelStart        =   1
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "1ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   1185
            TabIndex        =   185
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "AutoReveal Duration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   90
            LinkItem        =   "150"
            TabIndex        =   184
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   0
            Width           =   1965
         End
         Begin VB.Label lblBehaviourPopUpDelayMsCurrrent 
            Caption         =   "(0)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5100
            TabIndex        =   183
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   480
         End
         Begin VB.Label lblAutoRevealDurationMsHigh 
            Caption         =   "1000ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4455
            TabIndex        =   182
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   585
         End
      End
      Begin VB.Frame fraAutoHideDelay 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   435
         TabIndex        =   174
         Top             =   2970
         Width           =   6120
         Begin CCRSlider.Slider sliBehaviourAutoHideDelay 
            Height          =   315
            Left            =   1605
            TabIndex        =   175
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Max             =   2000
            TickFrequency   =   200
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "3s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   1245
            TabIndex        =   179
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   600
         End
         Begin VB.Label lblAutoHideDelayMsHigh 
            Caption         =   "5s"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4440
            TabIndex        =   178
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   405
         End
         Begin VB.Label lblAutoHideDelayMsCurrent 
            Caption         =   "(5) secs"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4950
            TabIndex        =   177
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "AutoHide Delay"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   105
            LinkItem        =   "150"
            TabIndex        =   176
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   -30
            Width           =   1350
         End
      End
      Begin VB.CheckBox chkBehaviourMouseActivate 
         Caption         =   "Pop-up on Mouseover"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4380
         TabIndex        =   173
         ToolTipText     =   "Essential functionality for the dock - pops up when  given focus"
         Top             =   8070
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame fraAnimationInterval 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   195
         TabIndex        =   153
         Top             =   6930
         Width           =   6180
         Begin CCRSlider.Slider sliAnimationInterval 
            Height          =   315
            Left            =   1890
            TabIndex        =   154
            ToolTipText     =   $"dockSettings.frx":130BD
            Top             =   285
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            Min             =   1
            Max             =   20
            Value           =   10
            TickFrequency   =   5
            SelStart        =   1
         End
         Begin VB.Label lblAnimationIntervalMsLow 
            Caption         =   "1ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1500
            TabIndex        =   158
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblAnimationIntervalMsHigh 
            Caption         =   "20ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4680
            TabIndex        =   157
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblAnimationIntervalMsCurrent 
            Caption         =   "(20)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5265
            TabIndex        =   156
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblBehaviourLabel 
            Caption         =   "Animation Interval"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   345
            LinkItem        =   "150"
            TabIndex        =   155
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   15
            Width           =   1605
         End
      End
      Begin VB.Frame fraIconEffect 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   105
         TabIndex        =   99
         Top             =   945
         Width           =   5025
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Sound Selection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   540
         TabIndex        =   220
         Top             =   6195
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Icon Origin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   540
         TabIndex        =   216
         ToolTipText     =   "The original icons may be low quality."
         Top             =   5670
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Lock the Dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   540
         TabIndex        =   214
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
         Top             =   5190
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Dock Hiding Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   525
         LinkItem        =   "150"
         TabIndex        =   203
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4545
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   $"dockSettings.frx":1314C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   12
         Left            =   1740
         TabIndex        =   170
         Top             =   7755
         Width           =   4485
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Style Themes and Fonts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   3
      Left            =   1665
      TabIndex        =   39
      ToolTipText     =   "This panel allows you to change the styling of the icon labels and the dock background image"
      Top             =   15
      Width           =   6930
      Begin VB.CheckBox chkLabelBackgrounds 
         Caption         =   "Enable Label Backgrounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3525
         TabIndex        =   171
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4065
         Width           =   2490
      End
      Begin VB.Frame fraFontOpacity 
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   210
         TabIndex        =   100
         ToolTipText     =   "The theme background "
         Top             =   6750
         Width           =   6525
         Begin CCRSlider.Slider sliStyleShadowOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   101
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   750
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin CCRSlider.Slider sliStyleOutlineOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   102
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1245
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin CCRSlider.Slider sliStyleFontOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   165
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   1635
            TabIndex        =   169
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   270
            Width           =   540
         End
         Begin VB.Label Label30 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4680
            TabIndex        =   168
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label lblStyleFontOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   167
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "Font Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   480
            TabIndex        =   166
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   -15
            Width           =   1350
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "Outline Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   450
            TabIndex        =   110
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   975
            Width           =   1365
         End
         Begin VB.Label lblStyleOutlineOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   109
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label Label35 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4665
            TabIndex        =   108
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   1635
            TabIndex        =   107
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "Shadow Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   465
            TabIndex        =   106
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label lblStyleShadowOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   105
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   765
            Width           =   630
         End
         Begin VB.Label Label39 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4680
            TabIndex        =   104
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   780
            Width           =   555
         End
         Begin VB.Label lblStyleLabel 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   1635
            TabIndex        =   103
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   780
            Width           =   540
         End
      End
      Begin VB.PictureBox picStylePreview 
         Height          =   735
         Left            =   630
         ScaleHeight     =   675
         ScaleWidth      =   5280
         TabIndex        =   51
         ToolTipText     =   $"dockSettings.frx":131DE
         Top             =   4440
         Width           =   5340
         Begin VB.Label lblPreviewFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2355
            TabIndex        =   52
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2400
            TabIndex        =   143
            Top             =   285
            Width           =   570
         End
         Begin VB.Label lblPreviewLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2340
            TabIndex        =   144
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblPreviewRight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2370
            TabIndex        =   145
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lblPreviewTop 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2355
            TabIndex        =   146
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblPreviewBottom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2355
            TabIndex        =   147
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2415
            TabIndex        =   148
            Top             =   285
            Width           =   570
         End
      End
      Begin VB.CommandButton btnStyleOutline 
         Caption         =   "&Outline Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "The colour of the outline, click the button to change"
         Top             =   6180
         Width           =   1470
      End
      Begin VB.CommandButton btnStyleShadow 
         Caption         =   "&Shadow Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5775
         Width           =   1470
      End
      Begin VB.CommandButton btnStyleFont 
         Caption         =   "Select &Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "The font used in the labels, click the button to change"
         Top             =   5370
         Width           =   1470
      End
      Begin VB.CheckBox chkStyleDisable 
         Caption         =   "Disable Icon Labels"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   630
         TabIndex        =   47
         ToolTipText     =   "You can toggle the icon labels on/off here"
         Top             =   4065
         Width           =   2235
      End
      Begin VB.ComboBox cmbStyleTheme 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":13268
         Left            =   2205
         List            =   "dockSettings.frx":1326A
         TabIndex        =   40
         ToolTipText     =   "The dock background theme can be selected here"
         Top             =   405
         Width           =   2520
      End
      Begin CCRSlider.Slider sliStyleOpacity 
         Height          =   315
         Left            =   2085
         TabIndex        =   42
         ToolTipText     =   "The theme background opacity is set here"
         Top             =   900
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Max             =   100
         TickFrequency   =   10
      End
      Begin CCRSlider.Slider sliStyleThemeSize 
         Height          =   315
         Left            =   2085
         TabIndex        =   159
         ToolTipText     =   "The theme background overall size is set here"
         Top             =   1335
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   1
         Max             =   177
         Value           =   30
         TickFrequency   =   10
         SelStart        =   50
      End
      Begin VB.Image imgThemeSample 
         BorderStyle     =   1  'Fixed Single
         Height          =   2070
         Left            =   645
         Picture         =   "dockSettings.frx":1326C
         Top             =   1800
         Width           =   5430
      End
      Begin VB.Label lblChkLabelBackgrounds 
         Caption         =   "Enable Label Backgrounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3795
         TabIndex        =   172
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblStyleLabel 
         Caption         =   "1px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1650
         TabIndex        =   160
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblStyleLabel 
         Caption         =   "Theme Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   660
         TabIndex        =   163
         ToolTipText     =   "The theme background overall size is set here"
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label lblStyleSizeCurrent 
         Caption         =   "(118px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   162
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblThemeSizeTextHigh 
         Caption         =   "118px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   161
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label Label999 
         Height          =   375
         Left            =   720
         TabIndex        =   142
         Top             =   7560
         Width           =   4215
      End
      Begin VB.Label lblStyleLabel 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1815
         TabIndex        =   46
         Top             =   945
         Width           =   420
      End
      Begin VB.Label lblStyleOutlineColourDesc 
         Caption         =   "Shadow Colour: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   70
         Top             =   6225
         Width           =   2700
      End
      Begin VB.Label lblStyleFontFontShadowColor 
         Caption         =   "Shadow Colour:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   69
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5820
         Width           =   2490
      End
      Begin VB.Label lblStyleFontOutlineTest 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         TabIndex        =   65
         ToolTipText     =   "The colour of the outline, click the button to change"
         Top             =   6225
         Width           =   390
      End
      Begin VB.Label lblStyleFontFontShadowTest 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         TabIndex        =   64
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5820
         Width           =   450
      End
      Begin VB.Label lblStyleFontName 
         Caption         =   "Font : Open Sans"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   63
         ToolTipText     =   "The font used in the labels, click the button to change"
         Top             =   5445
         Width           =   3765
      End
      Begin VB.Label Label44 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   45
         Top             =   945
         Width           =   585
      End
      Begin VB.Label lblStyleOpacityCurrent 
         Caption         =   "(0%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   44
         Top             =   945
         Width           =   630
      End
      Begin VB.Label lblStyleLabel 
         Caption         =   "Opacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   675
         TabIndex        =   43
         ToolTipText     =   "The theme background opacity is set here"
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label lblStyleLabel 
         Caption         =   "Theme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   675
         TabIndex        =   41
         ToolTipText     =   "The dock background theme can be selected here"
         Top             =   435
         Width           =   795
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Position the Dock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   4
      Left            =   1665
      TabIndex        =   22
      ToolTipText     =   "This panel controls the positioning of the whole dock"
      Top             =   30
      Width           =   6930
      Begin VB.CheckBox chkMoveWinTaskbar 
         Caption         =   "Avoid clashes with Windows taskbar?"
         Height          =   315
         Left            =   2205
         TabIndex        =   251
         Top             =   1755
         Width           =   3675
      End
      Begin VB.ComboBox cmbPositionLayering 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":16D96
         Left            =   2190
         List            =   "dockSettings.frx":16DA3
         TabIndex        =   37
         Text            =   "Always Below"
         ToolTipText     =   "Should the dock appear on top of other windows or underneath?"
         Top             =   2400
         Width           =   2595
      End
      Begin VB.ComboBox cmbPositionMonitor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":16DCC
         Left            =   2205
         List            =   "dockSettings.frx":16DE2
         TabIndex        =   26
         Text            =   "Monitor 1"
         ToolTipText     =   "Here you can determine upon which monitor the dock will appear"
         Top             =   480
         Width           =   2565
      End
      Begin VB.ComboBox cmbPositionScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":16E28
         Left            =   2190
         List            =   "dockSettings.frx":16E38
         TabIndex        =   25
         Text            =   "Bottom"
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1185
         Width           =   2595
      End
      Begin CCRSlider.Slider sliPositionEdgeOffset 
         Height          =   315
         Left            =   2085
         TabIndex        =   23
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3705
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   -15
         Max             =   128
         TickFrequency   =   8
      End
      Begin CCRSlider.Slider sliPositionCentre 
         Height          =   315
         Left            =   2085
         TabIndex        =   24
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   3075
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   -100
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.Label Label2 
         Caption         =   "Move Taskbar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         TabIndex        =   252
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Image imgMultipleGears3 
         Height          =   3000
         Left            =   3540
         Top             =   4980
         Width           =   2970
      End
      Begin VB.Image imgMultipleGears1 
         Height          =   3825
         Left            =   315
         Stretch         =   -1  'True
         Top             =   4485
         Width           =   2925
      End
      Begin VB.Label Label33 
         Caption         =   "Layering"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         TabIndex        =   38
         ToolTipText     =   "Should the dock appear on top of other windows or underneath?"
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label lblPositionMonitor 
         Caption         =   "Monitor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   36
         ToolTipText     =   "Here you can determine upon which monitor the dock will appear"
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label32 
         Caption         =   "Dock Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         TabIndex        =   35
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Centre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   34
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   3105
         Width           =   795
      End
      Begin VB.Label lblPositionCentrePercCurrent 
         Caption         =   "(0%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   33
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   3105
         Width           =   630
      End
      Begin VB.Label Label29 
         Caption         =   "+100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   32
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   3105
         Width           =   585
      End
      Begin VB.Label Label28 
         Caption         =   "-100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1590
         TabIndex        =   31
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   3105
         Width           =   630
      End
      Begin VB.Label Label27 
         Caption         =   "Edge Offset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   30
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3750
         Width           =   990
      End
      Begin VB.Label lblPositionEdgeOffsetPxCurrent 
         Caption         =   "(5px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   29
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3735
         Width           =   630
      End
      Begin VB.Label Label25 
         Caption         =   "128px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4890
         TabIndex        =   28
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3750
         Width           =   555
      End
      Begin VB.Label Label24 
         Caption         =   "-15px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1650
         TabIndex        =   27
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3750
         Width           =   540
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Icon Characteristics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   1
      Left            =   1665
      TabIndex        =   76
      ToolTipText     =   "This panel allows you to set the icon sizes and hover effects"
      Top             =   15
      Width           =   6930
      Begin VB.PictureBox picSizePreview 
         Height          =   4065
         Left            =   105
         ScaleHeight     =   4005
         ScaleWidth      =   6645
         TabIndex        =   136
         Top             =   4425
         Width           =   6705
         Begin VB.PictureBox picMinSize 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1920
            Left            =   0
            ScaleHeight     =   1920
            ScaleWidth      =   1920
            TabIndex        =   138
            ToolTipText     =   "The icon size in the dock when static"
            Top             =   915
            Width           =   1920
         End
         Begin VB.PictureBox picZoomSize 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3840
            Left            =   2775
            ScaleHeight     =   3840
            ScaleWidth      =   3840
            TabIndex        =   137
            ToolTipText     =   "The maximum icon size of an animated icon"
            Top             =   15
            Width           =   3840
         End
         Begin VB.Label Label1 
            Caption         =   "Icon Sizing Preview"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   141
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Icon size fully zoomed "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4575
            TabIndex        =   140
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   3810
            Width           =   1875
         End
         Begin VB.Label Label13 
            Caption         =   "Size of icon in the dock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   285
            TabIndex        =   139
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   3810
            Width           =   1875
         End
      End
      Begin VB.Frame fraZoomConfigs 
         BorderStyle     =   0  'None
         Height          =   1110
         Left            =   195
         TabIndex        =   111
         Top             =   3165
         Width           =   6495
         Begin CCRSlider.Slider sliIconsDuration 
            Height          =   315
            Left            =   1845
            TabIndex        =   112
            ToolTipText     =   "How long the effect is applied"
            Top             =   735
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   100
            Max             =   500
            Value           =   100
            TickFrequency   =   50
            SelStart        =   100
         End
         Begin CCRSlider.Slider sliIconsZoomWidth 
            Height          =   315
            Left            =   1845
            TabIndex        =   113
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   195
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   2
            Value           =   2
            SelStart        =   2
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "100ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   1320
            TabIndex        =   121
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   525
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "500ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   4650
            TabIndex        =   120
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   555
         End
         Begin VB.Label lblIconsDurationMsCurrent 
            Caption         =   "(200ms)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5265
            TabIndex        =   119
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "Duration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   480
            TabIndex        =   118
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   795
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   1665
            TabIndex        =   117
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label14 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4650
            TabIndex        =   116
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lblIconsZoomWidth 
            Caption         =   "(5)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5295
            TabIndex        =   115
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   225
            Width           =   630
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "Zoom Width"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   465
            TabIndex        =   114
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.CheckBox chkIconsZoomOpaque 
         Caption         =   "Zoom Opaque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   79
         ToolTipText     =   "Should the zoom be opaque too?"
         Top             =   1320
         Width           =   2685
      End
      Begin VB.ComboBox cmbIconsQuality 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":16E56
         Left            =   2160
         List            =   "dockSettings.frx":16E63
         TabIndex        =   78
         Text            =   "Low quality (Faster)"
         ToolTipText     =   $"dockSettings.frx":16EA5
         Top             =   390
         Width           =   2520
      End
      Begin CCRSlider.Slider sliIconsZoom 
         Height          =   315
         Left            =   2040
         TabIndex        =   77
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2775
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   1
         Max             =   256
         Value           =   1
         TickFrequency   =   32
         SelStart        =   1
      End
      Begin CCRSlider.Slider sliIconsSize 
         Height          =   315
         Left            =   2040
         TabIndex        =   80
         ToolTipText     =   "The size of each icon in the dock before any effect is applied"
         Top             =   2190
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   16
         Max             =   128
         Value           =   16
         TickFrequency   =   14
         SelStart        =   16
      End
      Begin CCRSlider.Slider sliIconsOpacity 
         Height          =   315
         Left            =   2040
         TabIndex        =   81
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   900
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         Min             =   50
         Max             =   100
         Value           =   50
         TickFrequency   =   7
         SelStart        =   50
      End
      Begin VB.Frame fraHoverEffect 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   150
         TabIndex        =   122
         Top             =   1575
         Width           =   6705
         Begin VB.ComboBox cmbIconsHoverFX 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":16F63
            Left            =   1995
            List            =   "dockSettings.frx":16F76
            TabIndex        =   123
            Text            =   "None"
            ToolTipText     =   "The zoom effect to apply"
            Top             =   105
            Width           =   2595
         End
         Begin VB.Label lblCharacteristicsLabel 
            Caption         =   "Hover Effect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   480
            TabIndex        =   124
            ToolTipText     =   "The zoom effect to apply"
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.Label lblchkIconsZoomOpaque 
         Caption         =   "Zoom Opaque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2415
         TabIndex        =   212
         Top             =   1305
         Width           =   2820
      End
      Begin VB.Label lblHidText3 
         Caption         =   "Some animation options are unavailable when running SteamyDock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         TabIndex        =   125
         Top             =   3585
         Width           =   5325
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "Quality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   630
         TabIndex        =   94
         ToolTipText     =   "Lower power machines will benefit from the lower quality setting"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "Opacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   630
         TabIndex        =   93
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   795
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "Icon Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   630
         TabIndex        =   92
         ToolTipText     =   "The size of each icon in the dock before any effect is applied"
         Top             =   2235
         Width           =   795
      End
      Begin VB.Label lblIconsOpacity 
         Caption         =   "(100%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   91
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label lblIconsSize 
         Caption         =   "(19px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   90
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4785
         TabIndex        =   89
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1710
         TabIndex        =   88
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "128px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4845
         TabIndex        =   87
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "16px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   1635
         TabIndex        =   86
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "Zoom Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   645
         TabIndex        =   85
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   795
      End
      Begin VB.Label lblIconsZoom 
         Caption         =   "(19px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   84
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   630
      End
      Begin VB.Label lblIconsZoomSizeMax 
         Caption         =   "256px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4845
         TabIndex        =   83
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   585
      End
      Begin VB.Label lblCharacteristicsLabel 
         Caption         =   "1px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   1755
         TabIndex        =   82
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   630
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "About SteamyDock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   6
      Left            =   1665
      TabIndex        =   54
      ToolTipText     =   "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
      Top             =   15
      Width           =   6930
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   6825
         TabIndex        =   219
         Top             =   2175
         Width           =   75
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   6570
         TabIndex        =   218
         Top             =   2175
         Width           =   330
      End
      Begin VB.TextBox lblAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6360
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   217
         Text            =   "dockSettings.frx":16FB6
         Top             =   2175
         Width           =   6660
      End
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1545
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs used by Rocketdock"
         Top             =   420
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   795
         Width           =   1470
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1170
         Width           =   1470
      End
      Begin VB.Label Label20 
         Caption         =   "(32bit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2985
         TabIndex        =   207
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "Windows XP, Vista, 7, 8 && 10 + ReactOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   206
         Top             =   1560
         Width           =   2955
      End
      Begin VB.Label Label10 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   205
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblPunklabsLink 
         BackStyle       =   0  'Transparent
         Caption         =   "                                                                                                                        "
         Height          =   225
         Index           =   0
         Left            =   2175
         MousePointer    =   1  'Arrow
         TabIndex        =   95
         Top             =   870
         Width           =   1710
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2175
         TabIndex        =   74
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1815
         TabIndex        =   73
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2535
         TabIndex        =   72
         Top             =   510
         Width           =   525
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2010
         TabIndex        =   71
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   62
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label Label60 
         Caption         =   "Dean Beedell � 2018"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   61
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label Label74 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   60
         Top             =   495
         Width           =   795
      End
      Begin VB.Label Label65 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   59
         Top             =   855
         Width           =   795
      End
      Begin VB.Label Label61 
         Caption         =   "Punklabs � 2005-2007"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   58
         Top             =   855
         Width           =   2175
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "General Configuration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   0
      Left            =   1665
      TabIndex        =   1
      ToolTipText     =   "These are the main settings for the dock"
      Top             =   15
      Width           =   6930
      Begin VB.Frame fraEditors 
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   825
         TabIndex        =   222
         Top             =   6660
         Width           =   5280
         Begin VB.CommandButton btnGeneralIconSettingsEditor 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   232
            Top             =   1260
            Width           =   300
         End
         Begin VB.CommandButton btnGeneralDockSettingsEditor 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   231
            Top             =   810
            Width           =   300
         End
         Begin VB.CommandButton btnGeneralDockEditor 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   230
            Top             =   330
            Width           =   300
         End
         Begin VB.TextBox txtIconSettingsDefaultEditor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   227
            Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
            Top             =   1275
            Width           =   3585
         End
         Begin VB.TextBox txtDockSettingsDefaultEditor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   225
            Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
            Top             =   795
            Width           =   3585
         End
         Begin VB.TextBox txtDockDefaultEditor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   223
            Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
            Top             =   345
            Width           =   3585
         End
         Begin VB.Label lblGenLabel 
            Caption         =   "VB6/TwinBasic Editor VBP or TwinProj file locations:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   229
            ToolTipText     =   $"dockSettings.frx":1772E
            Top             =   15
            Width           =   4755
         End
         Begin VB.Label lblGenLabel 
            Caption         =   "Icon Settings :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   228
            Top             =   1290
            Width           =   1530
         End
         Begin VB.Label lblGenLabel 
            Caption         =   "Dock Settings :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   226
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label lblGenLabel 
            Caption         =   "Dock :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   90
            TabIndex        =   224
            Top             =   390
            Width           =   1530
         End
      End
      Begin VB.CheckBox chkShowIconSettings 
         Caption         =   "Automatically display Icon Settings after adding an icon to the dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   945
         TabIndex        =   209
         Top             =   6315
         Width           =   5115
      End
      Begin VB.CheckBox chkSplashStatus 
         Caption         =   "Show Splash Screen at Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   164
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   6000
         Width           =   3735
      End
      Begin VB.Frame fraRunAppIndicators 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   450
         TabIndex        =   126
         Top             =   3210
         Width           =   5955
         Begin CCRSlider.Slider sliRunAppInterval 
            Height          =   315
            Left            =   1020
            TabIndex        =   127
            ToolTipText     =   "The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   540
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Min             =   5
            Max             =   65
            Value           =   5
            TickFrequency   =   3
            SelStart        =   15
            Transparent     =   -1  'True
         End
         Begin VB.Label lblGenRunAppInterval2 
            Caption         =   "5s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   795
            TabIndex        =   131
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   420
            Width           =   630
         End
         Begin VB.Label lblGenRunAppInterval3 
            Caption         =   "65s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3900
            TabIndex        =   130
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   420
            Width           =   585
         End
         Begin VB.Label lblGenRunAppIntervalCur 
            Caption         =   "(15 seconds)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4635
            TabIndex        =   129
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblGenLabel 
            Caption         =   "Running Application Check Interval"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   750
            LinkItem        =   "150"
            TabIndex        =   128
            ToolTipText     =   "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
            Top             =   105
            Width           =   3210
         End
      End
      Begin VB.CommandButton btnGeneralRdFolder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   5745
         TabIndex        =   97
         ToolTipText     =   "Select the folder location of Rocketdock here"
         Top             =   5475
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CheckBox chkShowRunning 
         Caption         =   "Running Application Indicators"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   11
         ToolTipText     =   $"dockSettings.frx":177F8
         Top             =   2880
         Width           =   2985
      End
      Begin VB.CheckBox chkGenDisableAnim 
         Caption         =   "Disable Minimise Animations"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3585
         TabIndex        =   10
         ToolTipText     =   "If you dislike the minimise animation, click this"
         Top             =   2505
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.CheckBox chkOpenRunning 
         Caption         =   "Open Running Application Instance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   12
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   4185
         Width           =   3030
      End
      Begin VB.TextBox txtAppPath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   915
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "C:\programs"
         ToolTipText     =   "This is the extrapolated location of the currently selected dock. This is for information only."
         Top             =   5460
         Width           =   4710
      End
      Begin VB.ComboBox cmbDefaultDock 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":17897
         Left            =   2085
         List            =   "dockSettings.frx":178A1
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Rocketdock"
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
         Top             =   4710
         Width           =   2310
      End
      Begin VB.CheckBox chkGenMin 
         Caption         =   "Minimise Windows to the Dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         TabIndex        =   9
         ToolTipText     =   "This allows running applications to appear in the dock"
         Top             =   2175
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.CheckBox chkStartupRun 
         Caption         =   "Run at Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   915
         TabIndex        =   2
         ToolTipText     =   "This will cause the current dock to run when Windows starts"
         Top             =   495
         Width           =   1440
      End
      Begin VB.Frame fraWriteOptionButtons 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   510
         TabIndex        =   151
         Top             =   1755
         Width           =   6165
         Begin VB.OptionButton optGeneralWriteConfig 
            Caption         =   "Write Settings to SteamyDock's Own Configuration Area"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   390
            TabIndex        =   152
            ToolTipText     =   $"dockSettings.frx":178BD
            Top             =   15
            Width           =   5325
         End
      End
      Begin VB.Frame fraReadOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   540
         TabIndex        =   132
         Top             =   690
         Width           =   6315
         Begin VB.OptionButton optGeneralReadSettings 
            Caption         =   "Read Settings from Rocketdock's portable SETTINGS.INI (single-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   135
            ToolTipText     =   "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
            Top             =   165
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralReadRegistry 
            Caption         =   "Read Settings from RocketDock's Registry (multi-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   134
            ToolTipText     =   $"dockSettings.frx":17952
            Top             =   465
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralReadConfig 
            Caption         =   "Read Settings From SteamyDock's Own Configuration Area (modern)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   133
            ToolTipText     =   $"dockSettings.frx":17A15
            Top             =   780
            Width           =   5565
         End
         Begin VB.Label lbloptGeneralReadConfig 
            Caption         =   "Read Settings From SteamyDock's Own Configuration Area (modern)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   630
            TabIndex        =   211
            Top             =   780
            Width           =   5115
         End
      End
      Begin VB.Label lblSquiggle 
         Caption         =   "-oOo-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2310
         TabIndex        =   233
         ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label lblGenLabel 
         Caption         =   "Dock Folder Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   915
         TabIndex        =   96
         ToolTipText     =   $"dockSettings.frx":17AAA
         Top             =   5190
         Width           =   1695
      End
      Begin VB.Label lblGenLabel 
         Caption         =   "Default Dock"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   915
         TabIndex        =   67
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
         Top             =   4740
         Width           =   1530
      End
   End
   Begin VB.Label lblDragCorner 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   8595
      TabIndex        =   234
      ToolTipText     =   "drag me"
      Top             =   9405
      Width           =   345
   End
   Begin VB.Label Label26 
      Caption         =   "Show Splash Screen at Startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2490
      TabIndex        =   208
      ToolTipText     =   "Show Splash Screen on Start-up"
      Top             =   8205
      Width           =   3870
   End
   Begin VB.Menu mnupopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About this utility"
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font selection for this utility"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with paypal"
         Index           =   2
      End
      Begin VB.Menu mnuSweets 
         Caption         =   "Donate some sweets/candy with Amazon"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuHelpPdf 
            Caption         =   "View Help (HTML)"
         End
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version"
         End
         Begin VB.Menu mnuMoreIcons 
            Caption         =   "Visit Deviantart to download some more Icons"
         End
         Begin VB.Menu mnuWidgets 
            Caption         =   "See the complementary steampunk widgets"
         End
         Begin VB.Menu mnuFacebook 
            Caption         =   "Chat about SteamyDock functionality on Facebook"
         End
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuBringToCentre 
         Caption         =   "Centre Program on Main Monitor"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuseparator1 
         Caption         =   ""
      End
      Begin VB.Menu mnuDevOptions 
         Caption         =   "Developer Options"
         Begin VB.Menu mnuAppFolder 
            Caption         =   "Reveal Program Location in Windows Explorer"
         End
         Begin VB.Menu mnuEditWidget 
            Caption         =   "Edit Program Using..."
         End
         Begin VB.Menu mnuDebug 
            Caption         =   "Turn Developer Options ON"
         End
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close this Program"
      End
   End
End
Attribute VB_Name = "dockSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : dockSettings
' Author    : beededea
' Date      : 30/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'
' Credits :
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           KillApp code from an unknown, untraceable source, possibly on MSN.
'           Registry reading code from ALLAPI.COM.
'           Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.
'
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
'           Krool on the VBForums for his impressive common control replacements
'           si_the_geek for his special folder code
'           Gary Beene        Get list of drive letters https://www.garybeene.com/code/visual%20basic145.htm
'
'           HanneSThEGreaT changeWallpaper https://forums.codeguru.com/showthread.php?497353-VB6-How-Do-I-Change-The-Windows-WallPaper
'
' NOTE - Do not END this program within the IDE as GDI will not release memory and usage will grow and grow
' ALWAYS use the QUIT option on the application right click menu.

' NOTE - When adding new slider controls remember to add them to the themeing menu option for light/high contrast

' NOTE - When building the binary, ensure that the ccrslider.ocx is in the docksettings folder
'         The manifest should be modified to incorporate the ocx

' SETTINGS: There are four settings files:
' o The first is the RD settings file SETTINGS.INI that only exists if RD is NOT using the registry

' NOTE: Rocketdock overwrites its own settings.ini when it closes meaning that we have to always work on a copy.
' In addition, when SD determines that RD is using the registry it extracts the data and creates a temporary copy of the settings file that we work on.
' In this manner we are always working on a .ini file in the same manner only writing it back to the registry when the user hits 'save & restart' or 'apply'.

' o The second is our tools copy of RD's settings file, we copy the original or create our own from RD's registry settings
' o The third is the settings file for this tool only, to store its own local preferences.

' origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file in program files.
' tmpSettingsFile = App.path & "\tmpSettings.ini" ' a temporary copy of the settings file that we work on.
' toolSettingsFile = SpecialFolder_AppData & <utilityName> "\settings.ini" the tool's own settings file.

' o The fourth settings file is the dockSettings.ini that sits in this location:
' C:\Users\<username>\AppData\Roaming\steamyDock\
'
' When the flag to write the 3rd settings file is set in the dock settings utility,
' we will write the rocketdock variable values to this file.
'
' docksettings.ini is partitioned as follows:
'
' [Software\SteamyDock\DockSettings] - the dockSettings tool writes here
' [Software\SteamyDock\IconSettings] - the iconSettings tool writes here
' [Software\SteamyDock\SteamyDockSettings] - the dock itself could write here but in reality it will most likely be in the areas above
'
' re: toolSettingsFile - The utilities read their own config files for their own personal set up in their own folders in appdata
' Settings.ini, this is just for local settings that concern only the utility, look and feel
'
' eg.
' C:\Users\<username>\AppData\Roaming\dockSettings\settings.ini
'
' toolSettingsFile - Dock - the following items are currently inserted into the toolSettingsFile for the dockSettings utility
'
' [Software\SteamyDockSettings]
' defaultStrength = 400
' defaultStyle = False
' defaultFont=Centurion Light SF

' toolSettingsFile - Icons - the following items are currently inserted into the toolSettingsFile for the iconSettings utility

' [Software\SteamyDockSettings]
' defaultFolderNodeKey=C:\Program Files (x86)\SteamyDock\iconSettings\my collection ' this could be moved to the docksettings.ini later
' rdMapState = Visible ' as could this
' defaultSize = 8
' defaultStrength = False
' defaultStyle = False
' Quality = 1
' defaultFont=Centurion Light SF
'
' The registry and the original settings.ini that Rocketdock provides for variable storage are
' left-overs from XP days when the registry storage was trendy and encouraged, the use of program files
' for the settings.ini was a left-over from the days before the registry when settings were stored locally
' in the program files folder with no folder security. This program allows access to these to retain
' compatibility with Rocketdock but offers the fourth storage option to be compatible with modern windows requirements.

' Separate labels for checkboxes
' the reason there is a separate label for certain checkboxes is due to the way that VB6 greys out checkbox labels using specific fonts
' causing them to appear 'crinkled'. When a discrete label is created that is unattached to the chkbox then it greys out correctly.
' When the main checkbox/radio button is disabled, its width is reduced and the associated label is made visible.
' Note that the balloon tooltips only function on the controls and not on the labels.

Option Explicit

'Simulate MouseEnter event to reset the icons on one frame
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

'API to test whether the user is running as an administrator account
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Private busyCounter As Integer
Private totalBusyMaximum As Integer

Private Const COLOR_BTNFACE As Long = 15

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long


' Flag for debug mode
Private mbDebugMode As Boolean  ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without

Public origSettingsFile As String

' module level balloon tooltip variables for comboBoxes ONLY.
Private gcmbBehaviourActivationFXBalloonTooltip As String
Private gcmbBehaviourAutoHideTypeBalloonTooltip As String
Private gcmbHidingKeyBalloonTooltip As String
Private gcmbBehaviourSoundSelectionBalloonTooltip As String
Private gcmbStyleThemeBalloonTooltip As String
Private gcmbPositionMonitorBalloonTooltip As String
Private gcmbPositionScreenBalloonTooltip As String
Private gcmbPositionLayeringBalloonTooltip As String
Private gcmbIconsQualityBalloonTooltip As String
Private gcmbIconsHoverFXBalloonTooltip As String
Private gcmbDefaultDockBalloonTooltip As String
Private gcmbWallpaperBalloonTooltip As String
Private gcmbWallpaperStyleBalloonTooltip As String
Private gcmbWallpaperTimerIntervalBalloonTooltip As String

'------------------------------------------------------ STARTS
' Constants and APIs to create and subclass the dragCorner
'Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Private Types for determining Form sizing

Private pvtLastFormHeight As Long
Private pvtFormResizedByDrag As Boolean

Private Const pvtCFormHeight As Long = 10110
Private Const pvtCFormWidth  As Long = 8910
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Private Types for determining whether the app is already DPI aware, most useful when operating within the IDE, stops "already DPI aware " messages.

Private Declare Function IsProcessDPIAware Lib "user32.dll" () As Boolean

Private Enum PROCESS_DPI_AWARENESS
    Process_DPI_Unaware = 0
    Process_System_DPI_Aware = 1
    Process_Per_Monitor_DPI_Aware = 2
End Enum
#If False Then
    Dim Process_DPI_Unaware, Process_System_DPI_Aware, Process_Per_Monitor_DPI_Aware
#End If

' this sets DPI awareness for the scope of this process, be it the binary or the IDE
Private Declare Function SetProcessDpiAwareness Lib "shcore.dll" (ByVal Value As PROCESS_DPI_AWARENESS) As Long

'------------------------------------------------------ ENDS

Private busyTimerCount As Integer
Private gblConstraintRatio As Double



'---------------------------------------------------------------------------------------
' Procedure : btnPreviousWallpaper_Click
' Author    : beededea
' Date      : 18/04/2025
' Purpose   : next wallpaper
'---------------------------------------------------------------------------------------
'
Private Sub btnPreviousWallpaper_Click()
    Dim cmbWallpaperIndex As Integer: cmbWallpaperIndex = 0
    
   On Error GoTo btnPreviousWallpaper_Click_Error

    cmbWallpaperIndex = cmbWallpaper.ListIndex
    cmbWallpaperIndex = cmbWallpaperIndex - 1
    
    If cmbWallpaperIndex <= 0 Then
        cmbWallpaperIndex = 0
    End If
    
    cmbWallpaper.ListIndex = cmbWallpaperIndex
    
    cmbWallpaper.SetFocus

   On Error GoTo 0
   Exit Sub

btnPreviousWallpaper_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPreviousWallpaper_Click of Form dockSettings"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnNextWallpaper_Click
' Author    : beededea
' Date      : 18/04/2025
' Purpose   : next wallpaper
'---------------------------------------------------------------------------------------
'
Private Sub btnNextWallpaper_Click()
    Dim cmbWallpaperIndex As Integer: cmbWallpaperIndex = 0
    
    On Error GoTo btnNextWallpaper_Click_Error

    cmbWallpaperIndex = cmbWallpaper.ListIndex
    cmbWallpaperIndex = cmbWallpaperIndex + 1
    
    If cmbWallpaperIndex >= cmbWallpaper.ListCount - 1 Then
        cmbWallpaperIndex = cmbWallpaper.ListCount - 1
    End If
    
    cmbWallpaper.ListIndex = cmbWallpaperIndex
    
    cmbWallpaper.SetFocus

   On Error GoTo 0
   Exit Sub

btnNextWallpaper_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnNextWallpaper_Click of Form dockSettings"
    
End Sub

' note: all buttons need to be style = graphical in order to theme by colour
Private Sub btnNextWallpaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If rDEnableBalloonTooltips = "1" Then CreateToolTip btnNextWallpaper.hWnd, "To select the next wallpaper click this button..", _
                  TTIconInfo, "Help on the next wallpaper button ", , , , True
End Sub




Private Sub btnPreviousWallpaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If rDEnableBalloonTooltips = "1" Then CreateToolTip btnPreviousWallpaper.hWnd, "To select the previous wallpaper click this button.", _
                  TTIconInfo, "Help on the previous wallpaper button ", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 24/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()
   On Error GoTo btnSave_Click_Error

    Call saveOrRestart(False)

   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkMoveWinTaskbar_Click
' Author    : beededea
' Date      : 10/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkMoveWinTaskbar_Click()
   On Error GoTo chkMoveWinTaskbar_Click_Error

    rDMoveWinTaskbar = CStr(chkMoveWinTaskbar.Value)

   On Error GoTo 0
   Exit Sub

chkMoveWinTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkMoveWinTaskbar_Click of Form dockSettings"
End Sub

Private Sub chkMoveWinTaskbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If rDEnableBalloonTooltips = "1" Then CreateToolTip chkMoveWinTaskbar.hWnd, "If this toggle is enabled, Steamydock will attempt to move the Windows taskbar to the opposite side, top to bottom &c when the two overlap. It will do this by restarting explorer. Note there will be some initial flickering whilst explorer restarts. If you don't want Steamydock to restart explorer just log out and in again.", _
                  TTIconInfo, "Help on the Avoid Clashes check box ", , , , True

End Sub









'---------------------------------------------------------------------------------------
' Procedure : fmeBehaviour_MouseMove
' Author    : beededea
' Date      : 21/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeBehaviour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeBehaviour_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the behaviour pane. Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
        titleText = "Help on the Behaviour Pane Button."
        CreateToolTip fmeBehaviour.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

fmeBehaviour_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeBehaviour_MouseMove of Form dockSettings"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : fmeStyle_MouseMove
' Author    : beededea
' Date      : 21/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeStyle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeStyle_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the style, themes and fonts pane. This is used to configure the label and font settings."
        titleText = "Help on the Style Themes and Fonts Pane Button."
        CreateToolTip fmeStyle.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeStyle_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeStyle_MouseMove of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fmeWallpaper_MouseMove
' Author    : beededea
' Date      : 25/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeWallpaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

   On Error GoTo fmeWallpaper_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the wallpaper pane. The wallpaper Panel allows you to select and apply a background image as the desktop wallpaper."
        titleText = "Help on the Wallpaper Pane Button."
        CreateToolTip fmeWallpaper.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeWallpaper_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeWallpaper_MouseMove of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Initialize
' Author    : beededea
' Date      : 28/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Initialize()
   On Error GoTo Form_Initialize_Error

    dockSettingsYPos = vbNullString
    dockSettingsXPos = vbNullString

   On Error GoTo 0
   Exit Sub

Form_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : Load the dockSettings form
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
    ' variables declared
    Dim NameProcess As String
    Dim AppExists As Boolean
    Dim answer As VbMsgBoxResult
    
    ' initial values assigned
    NameProcess = vbNullString
    AppExists = False
    answer = vbNo
    
    ' other variable assignments
    defaultDock = 0
    debugflg = 0
    startupFlg = True
    rdAppPath = vbNullString
    busyCounter = 1
    rDEnableBalloonTooltips = "1"
    
    sDDockSettingsDefaultEditor = vbNullString ' "E:\vb6\rocketdock\docksettings.vbp"
    gblSdIconSettingsDefaultEditor = vbNullString
    sDDockDefaultEditor = vbNullString
    gblRdDebugFlg = vbNullString
    pvtFormResizedByDrag = False

    mnupopmenu.Visible = False

    On Error GoTo Form_Load_Error
    
    If debugflg = 1 Then Debug.Print "%Form_Load"
    
    ' set the application to be DPI aware using the 'forbidden' API.
    If IsProcessDPIAware() = False Then Call setDPIAware
    
    ' initialise local vars
    gblFormPrimaryHeightTwips = vbNullString
    
    ' subclass controls that need additional functionality that VB6 does not provide (balloon tooltips on comboboxes)
    Call subClassControls
    
    ' obtain all drive names
    Call getAllDriveNames(sAllDrives)
                           
    'if the process already exists then kill it
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "docksettings.exe"
        checkAndKill NameProcess, False, False, False
        'MsgBox "You now have two instances of this utility running, they will conflict..."
    End If
        
    ' check the Windows version
    Call testWindowsVersion(classicThemeCapable)
    
    ' set form resizing variables
    Call setFormResizingVarsAndProperties
    
    ' the frames can jump about in the IDE during development, this just places them accurately at runtime
    Call placeFrames
    
    'load the about text
    Call loadAboutText
      
    ' get the location of this tool's settings file
    Call getToolSettingsFile
    
    'load the highlighted images onto the pressed icons
    Call loadHighlightedImages

    ' turn on the timer that tests every 10 secs whether the visual theme has changed
    ' only on those o/s versions that need it
    
    If classicThemeCapable = True Then
        dockSettings.mnuAuto.Caption = "Auto Theme Disable"
        dockSettings.themeTimer.Enabled = True
    Else
        dockSettings.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"
        dockSettings.themeTimer.Enabled = False
    End If

    ' check where rocketdock is installed
    Call checkRocketdockInstallation
        
    ' check where steamyDock is installed
    Call checkSteamyDockInstallation
    
    'update a filed with the installation details
    txtAppPath.Text = dockAppPath
    
    ' if both docks are installed we need to determine which is the default
    Call checkDefaultDock
    
    ' locate the icon settings tool ini file so we can read the editor VBP file
    Call locateiconSettingsToolFile
    
    'load the resizing image into a hidden picbox
    picHiddenPicture.Picture = LoadPicture(App.Path & "\resources\images\gpu-z-256.gif")
    
    'read the correct config location according to the default selection
    Call readDockConfiguration
    
    ' read the dock settings from the new configuration file  - currently barely used, it is all in above readDockConfiguration
    Call readSettingsFile
    
    If gblFormPrimaryHeightTwips = vbNullString Then gblFormPrimaryHeightTwips = CStr(gblStartFormHeight)

    ' RD can use the different monitors, SD cannot yet.
    Call GetMonitorCount
    
    ' read the local tool settings file and do some local things for the first and only time
    Call readAndSetUtilityFont
    
    ' display the version number on the general panel
    Call displayVersionNumber
    
    ' click on the panel that is set by default
    Call imgIcon_MouseDown_Event(Val(rDOptionsTabIndex) - 1)
    
    ' set the theme on startup
    Call setThemeSkin
    
    sdChkToggleDialogs = GetINISetting("Software\DockSettings", "sdChkToggleDialogs", toolSettingsFile)
    
    If sdChkToggleDialogs = vbNullString Then sdChkToggleDialogs = "1" ' validate
    If sdChkToggleDialogs = "1" Then ' set
        chkToggleDialogs.Value = 1
    Else
        chkToggleDialogs.Value = 0
    End If

    ' set the tooltips for the utility
    Call setToolTips
    
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties(dockSettings)
    
    ' various elements need to have their visibility and size modified prior to display
    Call makeVisibleFormElements
    
    ' sets other characteristics of the form and menus
    Call adjustMainControls
    
    ' save the initial anchor positions of ALL the controls on the form, TwinBasic has anchors but VB6 does not.
    Call saveControlSizes(dockSettings, gblFormControlPositions(), gblStartFormWidth, gblStartFormHeight)
        
    ' set the height of the whole form according to previously saved values but not higher than the screen size
    Call setFormHeight
    
    ' note: the final act in startup is the form_resize_event that is triggered by the subclassed WM_EXITSIZEMOVE when the form is finally revealed
    startupFlg = False ' now negate the startup flag

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form dockSettings"
     
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseCommonVars
' Author    : beededea
' Date      : 24/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialiseCommonVars()

   On Error GoTo initialiseCommonVars_Error

    rDGeneralReadConfig = vbNullString 'GeneralReadConfig", dockSettingsFile)
    rDGeneralWriteConfig = vbNullString 'GeneralWriteConfig", dockSettingsFile)
    rDRunAppInterval = vbNullString 'RunAppInterval", dockSettingsFile)
    rDDefaultDock = vbNullString 'DefaultDock", dockSettingsFile)
    rDAnimationInterval = vbNullString 'AnimationInterval", dockSettingsFile)
    rDSkinSize = vbNullString 'SkinSize", dockSettingsFile)
    sDSplashStatus = vbNullString 'SplashStatus", dockSettingsFile)
    sDShowIconSettings = vbNullString 'ShowIconSettings", dockSettingsFile) '' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock
    
    sDFontOpacity = vbNullString 'FontOpacity", settingsFile)
    sDAutoHideType = vbNullString 'AutoHideType", settingsFile)
    sDShowLblBacks = vbNullString 'ShowLblBacks", settingsFile) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
    sDContinuousHide = vbNullString 'ContinuousHide", settingsFile) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files ' nn
    sDBounceZone = vbNullString 'BounceZone", settingsFile) ' .05 DAEB 12/07/2021 common2.bas Add the BounceZone as a configurable variable.

    ' development
    sDDefaultEditor = vbNullString 'dockDefaultEditor", settingsFile)
    sDDebugFlg = vbNullString 'debugFlg", settingsFile)

    'RocketDock compatible settings only
    rDVersion = vbNullString 'Version", settingsFile)
    rDHotKeyToggle = vbNullString 'HotKey-Toggle", settingsFile)
            
    rDtheme = vbNullString 'Theme", settingsFile)
    rDWallpaper = vbNullString 'Wallpaper", settingsFile)
    rDWallpaperStyle = vbNullString 'WallpaperStyle", settingsFile)
    rDAutomaticWallpaperChange = vbNullString 'AutomaticWallpaperChange", settingsFile)
    rDWallpaperTimerIntervalIndex = vbNullString 'WallpaperTimerIntervalIndex", settingsFile)
    rDWallpaperTimerInterval = vbNullString 'WallpaperTimerInterval", settingsFile)
    rDWallpaperLastTimeChanged = vbNullString 'WallpaperLastTimeChanged", settingsFile)
    rDTaskbarLastTimeChanged = vbNullString 'TaskbarLastTimeChanged", settingsFile)
    
    rDMoveWinTaskbar = vbNullString 'MoveWinTaskbar", settingsFile)
    
    rDThemeOpacity = vbNullString 'ThemeOpacity", settingsFile)
    rDIconOpacity = vbNullString 'IconOpacity", settingsFile)
    rDFontSize = vbNullString 'FontSize", settingsFile)
    rDFontFlags = vbNullString 'FontFlags", settingsFile)
    rDFontName = vbNullString 'FontName", settingsFile)
    rDFontColor = vbNullString 'FontColor", settingsFile)
    rDFontCharSet = vbNullString 'FontCharSet", settingsFile)
    rDFontOutlineColor = vbNullString 'FontOutlineColor", settingsFile)
    rDFontOutlineOpacity = vbNullString 'FontOutlineOpacity", settingsFile)
    rDFontShadowColor = vbNullString 'FontShadowColor", settingsFile)
    rDFontShadowOpacity = vbNullString 'FontShadowOpacity", settingsFile)
    rDIconMin = vbNullString 'IconMin", settingsFile)
    rdIconMax = vbNullString 'IconMax", settingsFile)
    rDZoomWidth = vbNullString 'ZoomWidth", settingsFile)
    rDZoomTicks = vbNullString 'ZoomTicks", settingsFile)
    rDAutoHide = vbNullString 'AutoHide", settingsFile) '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
    rDAutoHideDuration = vbNullString 'AutoHideTicks", settingsFile)
    rDAutoHideDelay = vbNullString 'AutoHideDelay", settingsFile)
    rDPopupDelay = vbNullString 'PopupDelay", settingsFile)
    rDIconQuality = vbNullString 'IconQuality", settingsFile)
    rDLangID = vbNullString 'LangID", settingsFile)
    rDHideLabels = vbNullString 'HideLabels", settingsFile)
    rDZoomOpaque = vbNullString 'ZoomOpaque", settingsFile)
    rDLockIcons = vbNullString 'LockIcons", settingsFile)
    rDRetainIcons = vbNullString 'RetainIcons", settingsFile) ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    
    rDManageWindows = vbNullString 'ManageWindows", settingsFile)
    rDDisableMinAnimation = vbNullString 'DisableMinAnimation", settingsFile)
    rDShowRunning = vbNullString 'ShowRunning", settingsFile)
    rDOpenRunning = vbNullString 'OpenRunning", settingsFile)
    rDHoverFX = vbNullString 'HoverFX", settingsFile)
    rDzOrderMode = vbNullString 'zOrderMode", settingsFile)
    rDMouseActivate = vbNullString 'MouseActivate", settingsFile)
    rDIconActivationFX = vbNullString 'IconActivationFX", settingsFile)
    rDSoundSelection = vbNullString 'SoundSelection", settingsFile)
    
    rDMonitor = vbNullString 'Monitor", settingsFile)
    rDSide = vbNullString 'Side", settingsFile)
    rDOffset = vbNullString 'Offset", settingsFile)
    rDvOffset = vbNullString 'vOffset", settingsFile)
    rDOptionsTabIndex = vbNullString

   On Error GoTo 0
   Exit Sub

initialiseCommonVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseCommonVars of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : IMPORTANT: Called at every twip of resising, goodness knows what interval, we barely use this, instead we subclass and look for WM_EXITSIZEMOVE
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()

    ' this is here to avoid another resize when constraining the form height/width ratio in the Form_Resize_Event below - and thus avoiding a resizing of all controls.
    If gblDoNotResize = True Then
        gblDoNotResize = False
        Exit Sub
    End If
    
    ' this flags to the subclassing event that a manual resize of the form has been carried out
    pvtFormResizedByDrag = True
    
    ' only call this if the resize is done in code
    If InIDE Or gblFormResizedInCode = True Then
        Call Form_Resize_Event
    End If
                
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form dockSettings"
            Resume Next
          End If
    End With
''    If fmeMain(1).Visible = True Then
''        Call sliIconsSize_Change
''        Call sliIconsZoom_Change
''    End If


End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : What to do when unloading the main form
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Error
    
    Call thisForm_Unload
    
    If debugflg = 1 Then Debug.Print "%" & "Form_Unload"
    
   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form dockSettings"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : thisForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine called from several places
'---------------------------------------------------------------------------------------
'
Public Sub thisForm_Unload() ' name follows VB6 standard naming convention
    On Error GoTo Form_Unload_Error

    Call saveMainFormPosition

    Call DestroyToolTip ' destroys any current balloon tooltip
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisForm_Unload of Form dockSettings"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 forms
'---------------------------------------------------------------------------------------
'
Public Sub unloadAllForms(ByVal endItAll As Boolean)
    
    Dim ofrm As Form
    Dim NameProcess As String: NameProcess = vbNullString
    Dim fcount As Integer: fcount = 0
    Dim useloop As Integer: useloop = 0
       
    On Error GoTo unloadAllForms_Error
    
    ' the following unload commands take a while to complete resulting in a seeming-delay after a close, this .hide does away with that
    
    Me.Hide
              
    ' stop all VB6 timers in the main form
    
    themeTimer.Enabled = False
    repaintTimer.Enabled = False
    busyTimer.Enabled = False
    positionTimer.Enabled = False
    
    ' unload the native VB6 forms
    
    Unload about
    Unload frmMessage
    Unload licence
    'Unload dockSettings ' this will be unloaded at the end of the form_unload
    
    ' remove all variable references to each form in turn
    
    Set about = Nothing
    Set frmMessage = Nothing
    Set licence = Nothing
    Set dockSettings = Nothing
   
    On Error Resume Next
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unloadAllForms of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadHigherResFormImages
' Author    : beededea
' Date      : 18/06/2023
' Purpose   : load the images for the classic or high brightness themes
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResFormImages()
    
    On Error GoTo loadHigherResFormImages_Error
      
    If Me.WindowState = vbMinimized Then Exit Sub
        
    Call setFormIconImages
    
   On Error GoTo 0
   Exit Sub

loadHigherResFormImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResFormImages of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setFormIconImages
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : here we load the images for the icons at a set size, here we can test the size and load larger size icons if required.
'---------------------------------------------------------------------------------------
'
Private Sub setFormIconImages()
    
    On Error GoTo setFormIconImages_Error
    
    If Val(gblResizeRatio) < 1.25 Then
        imgIcon(0).Picture = LoadPicture(App.Path & "\resources\images\general.jpg")
        imgIcon(1).Picture = LoadPicture(App.Path & "\resources\images\icons.jpg")
        imgIcon(2).Picture = LoadPicture(App.Path & "\resources\images\behaviour.jpg")
        imgIcon(3).Picture = LoadPicture(App.Path & "\resources\images\style.jpg")
        imgIcon(4).Picture = LoadPicture(App.Path & "\resources\images\position.jpg")
        imgIcon(5).Picture = LoadPicture(App.Path & "\resources\images\wallpaper.jpg")
        imgIcon(6).Picture = LoadPicture(App.Path & "\resources\images\about.jpg")
    Else
        imgIcon(0).Picture = LoadPicture(App.Path & "\resources\images\general-128.jpg")
        imgIcon(1).Picture = LoadPicture(App.Path & "\resources\images\icons-128.jpg")
        imgIcon(2).Picture = LoadPicture(App.Path & "\resources\images\behaviour-128.jpg")
        imgIcon(3).Picture = LoadPicture(App.Path & "\resources\images\style-128.jpg")
        imgIcon(4).Picture = LoadPicture(App.Path & "\resources\images\position-128.jpg")
        imgIcon(5).Picture = LoadPicture(App.Path & "\resources\images\wallpaper-128.jpg")
        imgIcon(6).Picture = LoadPicture(App.Path & "\resources\images\about-128.jpg")
    End If
        
   On Error GoTo 0
   Exit Sub

setFormIconImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormIconImages of Form dockSettings"

End Sub








'
'---------------------------------------------------------------------------------------
' Procedure : setFormResizingVarsAndProperties
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set form and control resizing characteristics
'---------------------------------------------------------------------------------------
'
Private Sub setFormResizingVarsAndProperties()

   On Error GoTo setFormResizingVarsAndProperties_Error

    With lblDragCorner
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
      .Visible = False
    End With
    
    lblDragCorner.Visible = True
    
    pvtFormResizedByDrag = False
    gblDoNotResize = False

    gblResizeRatio = 1
    
    'adjust for windows 10 change in border size
    Call adjustWindows10FormSize

   On Error GoTo 0
   Exit Sub

setFormResizingVarsAndProperties_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormResizingVarsAndProperties of Form dockSettings"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : adjustWindows10FormSize
' Author    : Magic Ink
' Date      : 25/03/2025
' Purpose   : In Windows Vista and above the Width and Height are the size of the component, including the borders
'                ScaleWidth and ScaleHeight works together with ScaleLeft, ScaleTop and
'                ScaleMode to define the coordinate system for the component. By default,
'                ScaleTop and ScaleLeft are zero, and ScaleWidth and ScaleHeight are Width and Height minus the border,
'                in vbTwips (the default ScaleMode)
'
'                width         = full window + theme border
'                ScaleWidth    = window without any theme border
'
' NOTE: there is some more border resizing done within restoreSizableFormBorderStyle after re-enabling borderStyle = 2 using API
'---------------------------------------------------------------------------------------
'
Private Sub adjustWindows10FormSize()
    
    Dim desiredClientHeight As Long: desiredClientHeight = 0
    Dim desiredClientWidth As Long: desiredClientWidth = 0
    Dim windowBorderWidth As Long: windowBorderWidth = 0
    Dim windowBorderHeight As Long: windowBorderHeight = 0

    On Error GoTo adjustWindows10FormSize_Error
    
    If pvtBIsWinVistaOrGreater = True Then

        desiredClientHeight = pvtCFormHeight
        desiredClientWidth = pvtCFormWidth
        windowBorderWidth = (Me.Width - Me.ScaleWidth) / 2
        windowBorderHeight = (Me.Height - Me.ScaleHeight) / 4
        
        gblDoNotResize = True
        gblStartFormHeight = windowBorderHeight + desiredClientHeight
        'gblStartFormHeight = desiredClientHeight
        
        'MsgBox "2 adjustWindows10FormSize " & gblStartFormHeight
        Me.Height = gblStartFormHeight
        
        gblDoNotResize = True
        gblStartFormWidth = windowBorderWidth + desiredClientWidth
        'gblStartFormWidth = desiredClientWidth
        Me.Width = gblStartFormWidth
        
    Else
         gblStartFormHeight = desiredClientHeight
         gblStartFormWidth = desiredClientWidth
         
    End If
    
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
               Me.ScaleTop + Me.ScaleHeight
   
   On Error GoTo 0
   Exit Sub

adjustWindows10FormSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustWindows10FormSize of Form dockSettings"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnApplyWallpaper_MouseDown
' Author    : beededea
' Date      : 07/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnApplyWallpaper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wallpaperFullPath As String: wallpaperFullPath = vbNullString
   
    On Error GoTo btnApplyWallpaper_MouseDown_Error
    
    rDWallpaperStyle = cmbWallpaperStyle.List(cmbWallpaperStyle.ListIndex)
    
    ' save the last time the wallpaper changed
    rDWallpaperLastTimeChanged = CStr(Now())

    If rDWallpaper <> "none selected" Then
        PutINISetting "Software\SteamyDock\DockSettings", "Wallpaper", rDWallpaper, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "WallpaperStyle", rDWallpaperStyle, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "WallpaperLastTimeChanged", rDWallpaperLastTimeChanged, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "WallpaperTimerInterval", rDWallpaperTimerInterval, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "WallpaperTimerIntervalIndex", rDWallpaperTimerIntervalIndex, dockSettingsFile
        
        wallpaperFullPath = sdAppPath & "\wallpapers\" & rDWallpaper
        
        Call changeWallpaper(wallpaperFullPath, rDWallpaperStyle)
    End If
    
    cmbWallpaper.SetFocus

   On Error GoTo 0
   Exit Sub

btnApplyWallpaper_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnApplyWallpaper_MouseDown of Form dockSettings"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticWallpaperChange_Click
' Author    : beededea
' Date      : 07/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticWallpaperChange_Click()
   On Error GoTo chkAutomaticWallpaperChange_Click_Error

    If chkAutomaticWallpaperChange.Value = 1 Then
       lblWallpaper(3).Enabled = True
       lblWallpaper(4).Enabled = True
       cmbWallpaperTimerInterval.Enabled = True
    Else
       lblWallpaper(3).Enabled = False
       lblWallpaper(4).Enabled = False
       cmbWallpaperTimerInterval.Enabled = False
    End If

    rDAutomaticWallpaperChange = CStr(chkAutomaticWallpaperChange.Value)

   On Error GoTo 0
   Exit Sub

chkAutomaticWallpaperChange_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAutomaticWallpaperChange_Click of Form dockSettings"
End Sub

Private Sub chkAutomaticWallpaperChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkAutomaticWallpaperChange.hWnd, "This checkbox enables a timer in the dock that will change the desktop background on an interval you define.", _
                  TTIconInfo, "Help on the Apply Wallpaper Timer", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWallpaperTimerInterval_Click
' Author    : beededea
' Date      : 06/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbWallpaperTimerInterval_Click()
   On Error GoTo cmbWallpaperTimerInterval_Click_Error
   
   If startupFlg = True Then Exit Sub

    rDWallpaperTimerIntervalIndex = CStr(cmbWallpaperTimerInterval.ListIndex)

   On Error GoTo 0
   Exit Sub

cmbWallpaperTimerInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWallpaperTimerInterval_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : saveMainFormPosition
' Author    : beededea
' Date      : 04/08/2023
' Purpose   : called from several locations saves the form X,Y positions
'---------------------------------------------------------------------------------------
'
Public Sub saveMainFormPosition()

   On Error GoTo saveMainFormPosition_Error

    ' save the current X and y position of this form to allow repositioning when restarting
    dockSettingsXPos = dockSettings.Left
    dockSettingsYPos = dockSettings.top
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\DockSettings", "dockSettingsXPos", dockSettingsXPos, toolSettingsFile
    PutINISetting "Software\DockSettings", "dockSettingsYPos", dockSettingsYPos, toolSettingsFile
    
    On Error GoTo 0
   Exit Sub

saveMainFormPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveMainFormPosition of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize_Event
' Author    : beededea
' Date      : 30/05/2023
' Purpose   : Called mostly by WM_EXITSIZEMOVE, the subclassed (intercepted) message that indicates that the window has just been moved.
'             (and on a mouseUp during a bottom-right drag of the additional corner indicator). Also, in code as specifcally required with an indicator flag.
'             This prevents a resize occurring during every twip movement and the controls resizing themselves continuously.
'             They now only resize when the form resize has completed.
'
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize_Event()

    Dim currentFontSize As Single: currentFontSize = 0
    
    On Error GoTo Form_Resize_Event_Error
    
    ' When minimised and a resize is called then simply exit.
    If Me.WindowState = vbMinimized Then Exit Sub
             
    ' move the drag corner label along with the form's bottom right corner
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
           Me.ScaleTop + Me.ScaleHeight - (lblDragCorner.Height + 40)
    
    If pvtFormResizedByDrag = True Then

        ' maintain the aspect ratio, note: this change calls this routine again...
        dockSettings.Width = dockSettings.Height / gblConstraintRatio
        
        If gblSuppliedFontSize = vbNullString Then gblSuppliedFontSize = GetINISetting("Software\DockSettings", vbNullString, toolSettingsFile)
        currentFontSize = CSng(Val(gblSuppliedFontSize))
        
        'MsgBox "3 " & gblStartFormHeight
        ' resize all controls on the form
        Call resizeControls(Me, gblFormControlPositions(), gblStartFormWidth, gblStartFormHeight, currentFontSize)

        Call loadHigherResFormImages
        
    Else
        If Me.WindowState = 0 Then ' normal
            If pvtLastFormHeight <> 0 Then
               gblFormResizedInCode = True
               dockSettings.Height = pvtLastFormHeight
               
               ' lblHeight.Caption = "Form_Resize_Event 2 " & dockSettings.Height
            End If
        End If
    End If
            
    gblFormResizedInCode = False
    pvtFormResizedByDrag = False
    
    Call writeFormHeight
                
    On Error GoTo 0
    Exit Sub

Form_Resize_Event_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize_Event of Form dockSettings"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Moved
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Non VB6-standard event caught by subclassing and intercepting the WM_EXITSIZEMOVE (WM_MOVED) event
'---------------------------------------------------------------------------------------
'
Public Sub Form_Moved(sForm As String)

    On Error GoTo Form_Moved_Error
        
    'passing a form name as it allows us to potentially subclass another form's movement
    Select Case sForm
        Case "dockSettings"
            ' call a resize of all controls only when the form resize (by dragging) has completed (mouseUP)
            If pvtFormResizedByDrag = True Then
            
                'MsgBox gblDockSettingsFormOldHeight & " " & dockSettings.Height
            
                ' test the current form height and width, if the same then it is a form_moved and not a form_resize.
                If dockSettings.Height = gblDockSettingsFormOldHeight And dockSettings.Width = gblDockSettingsFormOldWidth Then
                    Exit Sub
                Else
                    gblDockSettingsFormOldHeight = dockSettings.Height
                    gblDockSettingsFormOldWidth = dockSettings.Width
                    
                    Call Form_Resize_Event
                    pvtFormResizedByDrag = False
                End If
                
            End If
            
        Case Else
    End Select
    
   On Error GoTo 0
   Exit Sub

Form_Moved_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Moved of Form dockSettings"
End Sub

    



Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnAboutDebugInfo.hWnd, "This is the debugging option.", _
                  TTIconInfo, "Help on the About Button", , , , True
End Sub

Private Sub btnSaveRestart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSaveRestart.hWnd, "Apply your recent changes to the settings and save them.", _
                  TTIconInfo, "Help on the Apply Button", , , , True
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnClose.hWnd, "Close the Dock Settings Utility.", _
                  TTIconInfo, "Help on the Close Button", , , , True
End Sub
Private Sub btnApplyWallpaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnApplyWallpaper.hWnd, "Display the selected wallpaper on the desktop. Don't forget to save and restart to make all the changes stick.", _
                  TTIconInfo, "Help on the Apply Wallpaper Button", , , , True
End Sub


Private Sub btnDefaults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnDefaults.hWnd, "Revert ALL settings to the defaults.", _
                  TTIconInfo, "Help on the Set Defaults Button", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnDonate.hWnd, "Opens a browser window and sends you to the donation page on Amazon.", _
                  TTIconInfo, "Help on the Donate Button", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnFacebook.hWnd, "This will link you to the Rocket/SteamyDock users Group.", _
                  TTIconInfo, "Help on the FaceBook Button", , , , True
End Sub

' ----------------------------------------------------------------
' Procedure Name: btnGeneralDockEditor_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 01/03/2024
' ----------------------------------------------------------------
Private Sub btnGeneralDockEditor_Click()
    
    Call selectDockVBPFile
    
    On Error GoTo 0
    Exit Sub

btnGeneralDockEditor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGeneralDockEditor_Click, line " & Erl & "."

End Sub

Private Sub btnGeneralDockSettingsEditor_Click()
    Call selectDockSettingsVBPFile(False)
End Sub

Private Sub btnGeneralIconSettingsEditor_Click()
    Call selectIconSettingsVBPFile
End Sub

Private Sub btnGeneralRdFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGeneralRdFolder.hWnd, "Press this button to select the folder location of Rocketdock here. ", _
                  TTIconInfo, "Help on selecting a folder.", , , , True

End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnHelp.hWnd, "This button open the tool's HTML help page in your browser.", _
                  TTIconInfo, "Help on the Help Button", , , , True
End Sub

Private Sub btnStyleFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleFont.hWnd, "This button gives the font selection box. Here you set the font as shown on the icon labels.", _
                  TTIconInfo, "Help on the Font Selection Button.", , , , True
End Sub

Private Sub btnStyleOutline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleOutline.hWnd, "The colour of the outline, click the button to change.", _
                  TTIconInfo, "Help on the Outline Colour Selection Button.", , , , True
End Sub

Private Sub btnStyleShadow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleShadow.hWnd, "The colour of the shadow, click the button to change.", _
                  TTIconInfo, "Help on the Shadow Colour Selection Button.", , , , True
End Sub

Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnFacebook.hWnd, "Here you can visit the update location where you can download new versions of the programs used by Rocketdock.", _
                  TTIconInfo, "Help on the Update Button", , , , True
End Sub

Private Sub chkAutoHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkAutoHide.hWnd, "This checkbox acts as a toggle. You can determine whether the dock will auto-hide or not and the type of hide that is implemented. using Rocketdock  only supports one type of hide and that is the slide type. Steamydock gives you an additional fade or an instant disappear. The latter is lighter on CPU usage whilst the former two are animated and require a little cpu during the transition.", _
                  TTIconInfo, "Help on the AutoHide Checkbox.", , , , True
End Sub
'
'Private Sub chkGenAlwaysAsk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenAlwaysAsk.hWnd, "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as set above. ", _
'                  TTIconInfo, "Help on Confirming which dock to use.", , , , True
'End Sub

Private Sub chkGenDisableAnim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenDisableAnim.hWnd, "If you dislike the minimise animation, click this. ", _
                  TTIconInfo, "Help on disabling the minimise animation.", , , , True
End Sub

Private Sub chkLockIcons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkLockIcons.hWnd, "This is an essential option that stops you accidentally deleting your dock icons, click it!. ", _
                  TTIconInfo, "Help on Dragging, dropping to or from the dock.", , , , True
                  
End Sub

Private Sub chkGenMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenMin.hWnd, "This option allows running applications to be minimised, appearing in the dock. Supported by Rocketdock only.", _
                  TTIconInfo, "Help on mimising apps to the dock.", , , , True
End Sub

Private Sub chkOpenRunning_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkOpenRunning.hWnd, "If you click on an icon that is already running then it can open it or fire up another instance. ", _
                  TTIconInfo, "Help on the Running Application Indicators.", , , , True
End Sub

Private Sub chkShowRunning_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkShowRunning.hWnd, "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency set below. ", _
                  TTIconInfo, "Help on Showing Running Applications .", , , , True
End Sub

Private Sub chkStartupRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkStartupRun.hWnd, "When this checkbox is ticked it will cause the selected dock to run when Windows starts. ", _
                  TTIconInfo, "Help on the Start with Windows Checkbox", , , , True
End Sub

Private Sub chkIconsZoomOpaque_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkIconsZoomOpaque.hWnd, "Should the zoomed icons be opaque when the others are transparent? Not yet implemented in Steamydock. ", _
                  TTIconInfo, "Help on the Zoom Opacity Checkbox", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkLabelBackgrounds_Click
' Author    : beededea
' Date      : 09/04/2025
' Purpose   : add a background to the icon titles in dock's drawtext function
'---------------------------------------------------------------------------------------
'
Private Sub chkLabelBackgrounds_Click()

   On Error GoTo chkLabelBackgrounds_Click_Error

   sDShowLblBacks = chkLabelBackgrounds.Value ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
    
   On Error GoTo 0
   Exit Sub

chkLabelBackgrounds_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkLabelBackgrounds_Click of Form dockSettings"
    
End Sub

Private Sub chkLabelBackgrounds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkLabelBackgrounds.hWnd, "With this checkbox you can toggle the icon label background on/off.", _
                  TTIconInfo, "Help on Label Background Disable.", , , , True
End Sub


' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
'---------------------------------------------------------------------------------------
' Procedure : chkRetainIcons_Click
' Author    : beededea
' Date      : 07/09/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkRetainIcons_Click()

    On Error GoTo chkRetainIcons_Click_Error

    rDRetainIcons = chkRetainIcons.Value

    On Error GoTo 0
    Exit Sub

chkRetainIcons_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRetainIcons_Click of Form dockSettings"
            Resume Next
          End If
    End With

End Sub

Private Sub chkRetainIcons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkRetainIcons.hWnd, "When you drag a program binary to the dock it can take an automatically selected icon or you can retain the embedded icon within the binary file. The automatically selected icon will come from our own collection. An embedded icon may well be good enough to display but be aware, older binaries use very small or low quality icons.", _
                  TTIconInfo, "Help on Retaining Original Icons.", , , , True
End Sub

Private Sub chkSplashStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkSplashStatus.hWnd, "When this checkbox is ticked the dock shows a Splash Screen on Start-up.", _
                  TTIconInfo, "Help on the Splash Screen Checkbox", , , , True
End Sub

Private Sub chkStyleDisable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkStyleDisable.hWnd, "This checkbox disables the labels that appear above the icon in the dock.", _
                  TTIconInfo, "Help on Label Disable.", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkToggleDialogs_Click
' Author    : beededea
' Date      : 28/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkToggleDialogs_Click()
    
    ' .70 DAEB 16/05/2022 rDIConConfig.frm Read the chkToggleDialogs value from a file and save the value for next time
   On Error GoTo chkToggleDialogs_Click_Error

    If chkToggleDialogs.Value = 0 Then
       sdChkToggleDialogs = "0"
    Else
       sdChkToggleDialogs = "1"
    End If
    
    PutINISetting "Software\DockSettings", "sdChkToggleDialogs", sdChkToggleDialogs, toolSettingsFile

    Call setToolTips

   On Error GoTo 0
   Exit Sub

chkToggleDialogs_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkToggleDialogs_Click of Form dockSettings"
End Sub

Private Sub chkToggleDialogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkToggleDialogs.hWnd, "This checkbox acts as a toggle to enable/disable the balloon tooltips.", _
                  TTIconInfo, "Help on the Ballooon Tooltip Toggle", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbHidingKey_Click ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
' Author    : beededea
' Date      : 26/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbHidingKey_Click()
   On Error GoTo cmbHidingKey_Click_Error

    rDHotKeyToggle = cmbHidingKey.Text

   On Error GoTo 0
   Exit Sub

cmbHidingKey_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbHidingKey_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeMain_MouseMove
' Author    : beededea
' Date      : 28/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString
    
   On Error GoTo fmeMain_MouseMove_Error

    descriptiveText = vbNullString
    titleText = vbNullString
    
    If rDEnableBalloonTooltips = "1" Then
        If Index = 0 Then
            descriptiveText = "Use this panel to configure the general options that apply to the whole dock program. "
            titleText = "Help on the General Pane."
        ElseIf Index = 1 Then
            descriptiveText = "Use this panel to configure the icon characteristics that apply only to the icons themselves. "
            titleText = "Help on the Icon Characteristics Pane."
        ElseIf Index = 2 Then
            descriptiveText = "Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
            titleText = "Help on the Behaviour Pane."
        ElseIf Index = 3 Then
            descriptiveText = "Use this panel to configure the label and font settings."
            titleText = "Help on the Style Themes and Fonts Pane."
        ElseIf Index = 4 Then
            descriptiveText = "This pane is used to control the location of the dock. "
            titleText = "Help on the Position Pane."
        ElseIf Index = 5 Then
            descriptiveText = "This pane is used to select and change the desktop background wallpaper. "
            titleText = "Help on the Wallpaper Pane."
        ElseIf Index = 6 Then
            ' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
            fraScrollbarCover.Visible = True
            descriptiveText = "The About Panel provides the version number of this utility, useful information when reporting a bug. The text below this gives due credit to Punk labs for being the originator of  and gives thanks to them for coming up with such a useful tool and also to Apple who created the original idea for this whole genre of docks. This pane also gives access to some useful utilities."
            titleText = "Help on the About Pane Button."
        End If
    End If

    CreateToolTip fmeMain(Index).hWnd, descriptiveText, TTIconInfo, titleText, , , , True

   On Error GoTo 0
   Exit Sub

fmeMain_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeMain_MouseMove of Form dockSettings"

End Sub


Private Sub fraAnimationInterval_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 16/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
   On Error GoTo btnHelp_Click_Error

    Call mnuHelpPdf_click
    
   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSplashStatus_Click
' Author    : beededea
' Date      : 01/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSplashStatus_Click()
   On Error GoTo chkSplashStatus_Click_Error

    If chkSplashStatus.Value = 1 Then
        sDSplashStatus = "1"
    Else
        sDSplashStatus = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkSplashStatus_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkSplashStatus_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAutoHideType_Click
' Author    : beededea
' Date      : 25/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbAutoHideType_Click()

   On Error GoTo cmbAutoHideType_Click_Error

    sDAutoHideType = cmbAutoHideType.ListIndex
    
    If cmbAutoHideType.ListIndex = 2 Then
        lblBehaviourLabel(2).Enabled = False
        lblBehaviourLabel(8).Enabled = False
        sliAutoHideDuration.Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblBehaviourLabel(3).Enabled = False
        lblBehaviourLabel(9).Enabled = False
        lblAutoRevealDurationMsHigh.Enabled = False
        sliBehaviourPopUpDelay.Enabled = False
        lblBehaviourPopUpDelayMsCurrrent.Enabled = False
        
    Else
        lblBehaviourLabel(2).Enabled = True
        lblBehaviourLabel(8).Enabled = True
        sliAutoHideDuration.Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblBehaviourLabel(3).Enabled = True
        lblBehaviourLabel(9).Enabled = True
        lblAutoRevealDurationMsHigh.Enabled = True
        sliBehaviourPopUpDelay.Enabled = True
        lblBehaviourPopUpDelayMsCurrrent.Enabled = True
       
    End If
    

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

cmbAutoHideType_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAutoHideType_Click of Form dockSettings"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 28/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeVisibleFormElements()

    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0

    On Error GoTo makeVisibleFormElements_Error
    
    imgIcon(0).Visible = True
    imgIcon(1).Visible = True
    imgIcon(2).Visible = True
    imgIcon(3).Visible = True
    imgIcon(4).Visible = True
    imgIcon(5).Visible = True
    imgIcon(6).Visible = True

    imgIconPressed(0).Visible = False
    imgIconPressed(1).Visible = False
    imgIconPressed(2).Visible = False
    imgIconPressed(3).Visible = False
    imgIconPressed(4).Visible = False
    imgIconPressed(5).Visible = False
    imgIconPressed(6).Visible = False
    
    screenHeightTwips = GetDeviceCaps(Me.hDC, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(Me.hDC, HORZRES) * screenTwipsPerPixelX ' replaces buggy screen.width

    ' read the form X/Y params from the toolSettings.ini
'    dockSettingsYPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsYPos", toolSettingsFile)
'    dockSettingsXPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsXPos", toolSettingsFile)
'
'    If dockSettingsYPos <> "" Then
'        dockSettings.Top = Val(dockSettingsYPos)
'    Else
'        dockSettings.Top = Screen.Height / 2 - dockSettings.Height / 2
'    End If
'
'    If dockSettingsXPos <> "" Then
'        dockSettings.Left = Val(dockSettingsXPos)
'    Else
'        dockSettings.Left = Screen.Width / 2 - dockSettings.Width / 2
'    End If

    ' read the form's saved X/Y params from the toolSettings.ini in twips and convert to pixels
    formLeftPixels = Val(GetINISetting("Software\DockSettings", "dockSettingsXPos", toolSettingsFile)) / screenTwipsPerPixelX
    formTopPixels = Val(GetINISetting("Software\DockSettings", "dockSettingsYPos", toolSettingsFile)) / screenTwipsPerPixelY

    Call adjustFormPositionToCorrectMonitor(Me.hWnd, formLeftPixels, formTopPixels)

   On Error GoTo 0
   Exit Sub

makeVisibleFormElements_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeVisibleFormElements of Form dockSettings"
        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   If debugflg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    mnuDebug_Click

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuSweets_Click

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : busyTimer_Timer
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : rotates the hourglass timer
'---------------------------------------------------------------------------------------
'
Private Sub busyTimer_Timer()
        Dim thisWindow As Long: thisWindow = 0
        Dim busyFilename As String: busyFilename = vbNullString
        Static totalBusyCounter As Integer
        
        On Error GoTo busyTimer_Timer_Error

        thisWindow = FindWindowHandle("SteamyDock")
        busyFilename = vbNullString
        
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        totalBusyCounter = totalBusyCounter + 1
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        If classicTheme = True Then
            busyFilename = App.Path & "\resources\images\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.Path & "\resources\images\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        
        If thisWindow <> 0 And totalBusyCounter >= totalBusyMaximum Then
            busyTimer.Enabled = False
            busyCounter = 1
            totalBusyCounter = 1
            picBusy.Visible = False
        End If

   On Error GoTo 0
   Exit Sub

busyTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure busyTimer_Timer of Form dockSettings"

End Sub


Private Sub fraAutoHideDelay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHideDuration_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub




' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkShowIconSettings_Click
' Author    : beededea
' Date      : 01/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkShowIconSettings_Click()

    On Error GoTo chkShowIconSettings_Click_Error
    
    If chkShowIconSettings.Value = 1 Then
        sDShowIconSettings = "1"
    Else
        sDShowIconSettings = "0"
    End If
    
    On Error GoTo 0
    Exit Sub

chkShowIconSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowIconSettings_Click of Form dockSettings"
End Sub



Private Sub chkShowIconSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkShowIconSettings.hWnd, "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here.", _
                  TTIconInfo, "Help on the automatic icon Settings Startup", , , , True
End Sub



Private Sub imgWallpaperPreview_Click()
    cmbWallpaper.SetFocus
End Sub

Private Sub fmeWallpaperPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If rDEnableBalloonTooltips = "1" Then CreateToolTip fmeWallpaperPreview.hWnd, "This image box displays a resized preview version of a much larger wallpaper, press the change button to apply it to your desktop.", _
                  TTIconInfo, "Help on the Wallpaper Preview", , , , True

End Sub

' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
Private Sub lblAboutText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = False
End Sub

Private Sub lblChkLabelBackgrounds_Click()
' the reason there is a separate label for certain checkboxes is due to the way that VB6 greys out checkbox labels using specific fonts causing them to be crinkled. When the label is unattached to the chkbox then it greys out correctly.
    Call chkLabelBackgrounds_Click
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblGeneralWriteConfig_Click
' Author    : beededea
' Date      : 30/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGeneralWriteConfig_Click()

   On Error GoTo lblGeneralWriteConfig_Click_Error

    If optGeneralWriteConfig.Value = False Then
        optGeneralWriteConfig.Value = True
    End If

   On Error GoTo 0
   Exit Sub

lblGeneralWriteConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGeneralWriteConfig_Click of Form dockSettings"

End Sub


 


'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseMove
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo lblDragCorner_MouseMove_Error

    lblDragCorner.MousePointer = 8

    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form dockSettings"
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseDown
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo lblDragCorner_MouseDown_Error
    
    If Button = vbLeftButton Then
        pvtFormResizedByDrag = True
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAppFolder_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAppFolder_Click()
    Dim folderPath As String: folderPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuAppFolder_Click_Error

    folderPath = App.Path
    If fDirExists(folderPath) Then ' if it is a folder already

        execStatus = ShellExecute(Me.hWnd, "open", folderPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        MsgBox "Having a bit of a problem opening a folder for this widget - " & folderPath & " It doesn't seem to have a valid working directory set.", "Dock Settings Confirmation Message", vbOKOnly + vbExclamation
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "Dock Settings Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuAppFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAppFolder_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuEditWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuEditWidget_Click()

    On Error GoTo mnuEditWidget_Click_Error
    
    Call runDockSettingsVBPFile

   On Error GoTo 0
   Exit Sub

mnuEditWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuEditWidget_Click of Form menuForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadConfig_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : set a value to a flag that indicates we will read from the 3rd settings file
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadConfig_Click()


   On Error GoTo optGeneralReadConfig_Click_Error

    If startupFlg = True Then '
        ' don't do this on the first startup run
        Exit Sub

    End If
    
    If chkShowRunning.Value = 1 Then
'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
'    If optGeneralReadConfig.Value = True And defaultDock = 1 And steamyDockInstalled = True And rocketDockInstalled = True Then
'        'chkGenAlwaysAsk.Enabled = True
'        'lblChkAlwaysConfirm.Enabled = True
'    End If
    
    rDGeneralReadConfig = CStr(optGeneralReadConfig.Value) ' this is the nub
    
    'Call locateDockSettingsFile

   On Error GoTo 0
   Exit Sub

optGeneralReadConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadConfig_Click of Form dockSettings"
End Sub

Private Sub optGeneralReadConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadConfig.hWnd, "This stores ALL SteamyDock's configuration within the user data area. This option retains future compatibility within modern versions of Windows. Not applicable for Rocketdock ", _
                  TTIconInfo, "Help on using SteamyDock's config.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadRegistry_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadRegistry_Click()
   On Error GoTo optGeneralReadRegistry_Click_Error

        If optGeneralReadRegistry.Value = True Then
            ' nothing to do, the checkbox value is used later to determine where to write the data
        End If
        'If defaultDock = 0 Then optGeneralWriteRegistry.Value = True ' if running Rocketdock the two must be kept in sync
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
        'chkGenAlwaysAsk.Enabled = False
        'lblChkAlwaysConfirm.Enabled = False
        
        rDGeneralReadConfig = CStr(optGeneralReadConfig.Value) ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralReadRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadRegistry_Click of Form dockSettings"

End Sub

Private Sub optGeneralReadRegistry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadRegistry.hWnd, "This option allows you to read Rocketdock's configuration from the Rocketdock portion of the Registry. This method is becoming increasingly incompatible with newer Windows beyond XP as it can cause some security problems on newer system as it requires admin rights to write back. Use it here in a read-only fashion to migrate from Rocketdock.", _
                  TTIconInfo, "Help on reading from the registry", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadSettings_Click
' Author    : beededea
' Date      : 04/03/2020
' Purpose   : The existence of a file in the rocketdock program files location is the sole flag that Rocketdocki
'             uses to determine whether it should write the settings to the registry or the settings file.
'             The file is settings.ini.
'
'             I had expected a flag in the registry but none exists... When I created a file in the
'             Rocketdock folder and it used it straight away.
'
'             The changes only come into effect on a click of the 'apply' button.
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadSettings_Click()

   On Error GoTo optGeneralReadSettings_Click_Error
   If debugflg = 1 Then Debug.Print "%optGeneralReadSettings_Click"
   
    tmpSettingsFile = rdAppPath & "\tmpSettings.ini" ' temporary copy of Rocketdock 's settings file
    
    If startupFlg = True Then '
        ' don't do this on the first startup run
        Exit Sub
    Else
'        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
'            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
'            ' create a settings.ini file in the rocketdock folder
'            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
'            Close #1         ' close the file and
'             ' test it exists
'            If fFExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
'                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
'                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
'            End If
'        End If
    End If
        
'    lblGenLabel(0).Enabled = False
'    lblGenLabel(1).Enabled = False
    sliRunAppInterval.Enabled = False
    lblGenLabel(2).Enabled = False
    lblGenRunAppIntervalCur.Enabled = False
    'chkGenAlwaysAsk.Enabled = False
    'lblChkAlwaysConfirm.Enabled = False
    
    rDGeneralReadConfig = CStr(optGeneralReadConfig.Value) ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralReadSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadSettings_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readIconsWriteSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the registry icon store one line at a time and create a temporary settings file
'---------------------------------------------------------------------------------------
'
Private Sub readIconsWriteSettings(location As String, settingsFile As String)
    
    ' variables declared
    Dim useloop As Integer
    Dim regRocketdockSection As String
    Dim theCount As Integer
    
    ' initial values assigned
     useloop = 0
     regRocketdockSection = vbNullString
     theCount = 0
    
    On Error GoTo readIconsWriteSettings_Error
    If debugflg = 1 Then Debug.Print "%" & "readIconsWriteSettings"
    
    'initialise the dimensioned variables
    useloop = 0
    regRocketdockSection = vbNullString
    theCount = 0
        
    ' get items from the registry that are required to 'default' the dock but aren't controlled by the dock settings utility
    theCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count") 'dean
    rdIconCount = theCount - 1
            
    ' first we read and write the individual icon data values
    For useloop = 0 To rdIconCount
         ' get the relevant entries from the registry
         readRegistryIconValues (useloop)
         ' read the rocketdock alternative settings.ini
         Call writeIconSettingsIni(location & "\Icons", useloop, settingsFile) ' the alternative settings.ini exists when RD is set to use it
     Next useloop
     
    PutINISetting location & "\Icons", "Count", theCount, settingsFile
    
   On Error GoTo 0
   Exit Sub

readIconsWriteSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconsWriteSettings of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : readRegistryIconValues
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read the registry and set obtain the necessary icon data for the specific icon
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryIconValues(ByVal iconNumberToRead As Integer)
    ' read the settings from the registry
   On Error GoTo readRegistryIconValues_Error
   If debugflg = 1 Then Debug.Print "%" & "readRegistryIconValues"
  
    sFilename = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName")
    sFileName2 = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName2")
    sTitle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Title")
    sCommand = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Command")
    sArguments = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Arguments")
    sWorkingDirectory = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-WorkingDirectory")
    sShowCmd = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-ShowCmd")
    sOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-OpenRunning")
    sIsSeparator = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-IsSeparator")
    sUseContext = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseContext")
    sDockletFile = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-DockletFile")

   On Error GoTo 0
   Exit Sub

readRegistryIconValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryIconValues of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : fmeSizePreview_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeSizePreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo fmeSizePreview_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%fmeSizePreview_MouseDown"
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

fmeSizePreview_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeSizePreview_MouseDown of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkDefaultDock
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : if both rocketdock and steamydock are installed, then asks which dock you would like to maintain/configure
'---------------------------------------------------------------------------------------
'
Private Sub checkDefaultDock()

    ' variables declared
    Dim answer As VbMsgBoxResult
        
    ' initial values assigned
     answer = vbNo
    
   On Error GoTo checkDefaultDock_Error
   
    'initialise the dimensioned variables
    answer = vbNo
    
    If steamyDockInstalled = True Then
        ' get the location of the dock's new settings file
        Call locateDockSettingsFile
        'chkGenAlwaysAsk.Value = Val(GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile))
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
        If rDGeneralReadConfig <> vbNullString Then
            optGeneralReadConfig.Value = CBool(rDGeneralReadConfig)
        Else
            optGeneralReadConfig.Value = False
        End If
    End If
    
    'cmbDefaultDock.ListIndex = 1
    
    dockAppPath = sdAppPath
    txtAppPath.Text = dockAppPath
    defaultDock = 1
    ' write the default dock to the SteamyDock settings file
    PutINISetting "Software\DockSettings", "defaultDock", defaultDock, toolSettingsFile
                
'    If steamyDockInstalled = True And rocketDockInstalled = True Then
'        If chkGenAlwaysAsk.Value = 1 Then  ' depends upon being able to read the new configuration file in the user data area
'            answer = MsgBox("Both Rocketdock and SteamyDock are installed on this system. Use SteamyDock by default? ", vbYesNo)
'            If answer = vbYes Then
'                'cmbDefaultDock.ListIndex = 1 ' steamy dock
'                dockAppPath = sdAppPath
'                txtAppPath.Text = sdAppPath
'                defaultDock = 1
'            Else
'                'cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                txtAppPath.Text = rdAppPath
'                defaultDock = 0
'            End If
'        Else
'            ' if the question is not being asked then use the default dock as specified in the docksettings.ini file
'            If rDDefaultDock = "steamydock" Then
'                'cmbDefaultDock.ListIndex = 1
'                dockAppPath = sdAppPath
'                txtAppPath.Text = dockAppPath
'                defaultDock = 1
'            ElseIf rDDefaultDock = "rocketdock" Then
'                'cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                txtAppPath.Text = rdAppPath
'                defaultDock = 0
'            Else
''                If cmbDefaultDock.ListIndex = 1 Then  ' depends upon being able to read the new configuration file in the user data area
''                    dockAppPath = sdAppPath
''                    txtAppPath.Text = dockAppPath
''                    defaultDock = 1
''                Else
''                    cmbDefaultDock.ListIndex = 0 ' rocket dock
''                    dockAppPath = rdAppPath
''                    txtAppPath.Text = rdAppPath
''                    defaultDock = 0
''                End If
'            End If
'        End If
'    ElseIf steamyDockInstalled = True Then ' just steamydock installed
'            cmbDefaultDock.ListIndex = 1
'            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
'
'            dockAppPath = sdAppPath
'            txtAppPath.Text = dockAppPath
'            defaultDock = 1
'            ' write the default dock to the SteamyDock settings file
'            PutINISetting "Software\SteamyDockSettings", "defaultDock", defaultDock, toolSettingsFile
'
'    ElseIf rocketDockInstalled = True Then ' just rocketdock installed
'            cmbDefaultDock.ListIndex = 0
'            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
'
'            dockAppPath = rdAppPath
'            txtAppPath.Text = rdAppPath
'            defaultDock = 0
'    End If
    
    ' it is possible to run this program without steamydock being installed
    If steamyDockInstalled = False And rocketDockInstalled = False Then
        answer = MsgBox(" Neither Rocketdock nor SteamyDock has been installed on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
         Dim ofrm As Form
         For Each ofrm In Forms
             Unload ofrm
         Next
         End
    End If

   On Error GoTo 0
   Exit Sub

checkDefaultDock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkDefaultDock of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readDockConfiguration
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : read the configurations, settings.ini, registry and dockSettings.ini
'---------------------------------------------------------------------------------------
'
Private Sub readDockConfiguration()
    ' select the settings source STARTS
            
   On Error GoTo readDockConfiguration_Error

    'final check to be sure that we aren't using an incorrectly set dockSettings.ini file when RD has never been installed
    If rocketDockInstalled = False And RDregistryPresent = False Then
        rDGeneralReadConfig = True
        optGeneralReadConfig.Value = True
    End If

    If steamyDockInstalled = True And defaultDock = 1 And optGeneralReadConfig.Value = True Then ' it will always exist even if not used
        ' read the dock settings from the new configuration file
        Call initialiseCommonVars
        Call readDockSettingsFile("Software\SteamyDock\DockSettings", dockSettingsFile)
        Call validateInputs
        Call adjustControls

        rDVersion = App.Major & "." & App.Minor & "." & App.Revision

    End If
    
    If optGeneralReadConfig.Value = False And rocketDockInstalled = True Then
        ' read the dock settings from INI or from registry
        Call readRocketdockSettings
        Call adjustControls
    End If
    
    'if rocketdock set the automatic startup string to Steamydock
    rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SteamyDock")
    If rdStartupRunString <> vbNullString Then
        rDStartupRun = "1"
        chkStartupRun.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

readDockConfiguration_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockConfiguration of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadAboutText
' Author    : beededea
' Date      : 12/03/2020
' Purpose   : The text for the abour page is stored here
'---------------------------------------------------------------------------------------
'
Sub loadAboutText()
    On Error GoTo loadAboutText_Error
    If debugflg = 1 Then Debug.Print "%loadAboutText"
    
'    Dim strLine As String
'    Dim strFile As String
'    Dim intFile As Integer
'
'    strLine = ""
'    strFile = App.Path & "about.txt"
'    '
'    ' If the file exists then read it and
'    ' populate the TextBox
'    '
'    If fFExists(strFile) <> "" Then
'        intFile = FreeFile
'        Open strFile For Input As intFile
'        Line Input #1, strLine
'        lblAboutText.Text = strLine
'        Close intFile
'    End If


    Call LoadFileToTB(lblAboutText, App.Path & "\about.txt", False)

'    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project to allow easier configuration, bug-fixing and enhancement of Rocketdock and currently underway, a fully open source version of a Rocketdock clone."
'
'    lblAboutPara4.Caption = "The first steps are the two VB6 utilities that replicate the icons settings and dock settings screen. The subsequent step is the dock itself. I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received."
'
'    lblAboutPara1.Caption = "The original Rocketdock was developed by the Apple fanboy and fangirl team at Punklabs. They developed it as a peace offering from the Mac community to the Windows Community."
'    lblAboutPara2.Caption = "This new dock, now known as SteamyDock, was developed by a Windows/ReactOS fanboy on Windows 7 using VB6. This utility faithfully reproduces the original as created by Punklabs, originally done solely as a homage to the original as that version is no longer being supported but now it has evolved into a set of tools that has become a replacement for rocketdock itself. It must be said, the initial idea for this dock came from Punklabs and Rocketdock's OS/X dock predecessors. All HAIL to Punklabs!"
'    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project. It is open source to allow easier configuration, bug-fixing and enhancement of Rocketdock and community contribution towards this new dock."
'    lblAboutPara4.Caption = "The first steps were the two VB6 utilities that replicate the icons settings and dock settings screen (this utility). These are largely complete and the dock itself is now under development and 90% complete. A future step is conversion to RADBasic/TwinBasic or VB.NET for future-proofing and 64bit-ness. This next step is 1/3rd underway."
'
'    lblAboutPara5.Caption = "I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received if you choose to contribute."
'    lblAboutPara6.Caption = "This utility MUST run as administrator in order to access Rocketdock's " & _
'                            "registry settings (due to a Windows shadow registry feature/bug that " & _
'                            "gives incorrect shadow data). If you run it without admin rights and " & _
'                            "you want to change the values in the registry then some of the values may " & _
'                            "be incorrect and the resulting dock might look and act rather strange. " & _
'                            "You have been warned!"

   On Error GoTo 0
   Exit Sub

loadAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAboutText of Form dockSettings"
    
End Sub
    
    





   
'---------------------------------------------------------------------------------------
' Procedure : InIDE
' Author    : beededea
' Date      : 03/03/2020
' Purpose   : There are ocasions when the program will act differently when running in the IDE
'             We need to know when.  Compatibility mode means that it believes it is running under XP and will return as such.
'---------------------------------------------------------------------------------------
'
Function InIDE() As Boolean
'Returns whether we are running in vb(true), or compiled (false)
 
   On Error GoTo InIDE_Error
   If debugflg = 1 Then Debug.Print "%InIDE"

    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        InIDE = True
    End If

   On Error GoTo 0
   Exit Function

InIDE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InIDE of Form dockSettings"
 
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : InDebugMode
' Author    : beededea
' Date      : 02/03/2021
' Purpose   : ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
'---------------------------------------------------------------------------------------
'
Private Function InDebugMode() As Boolean
   On Error GoTo InDebugMode_Error

    mbDebugMode = True
    InDebugMode = True

   On Error GoTo 0
   Exit Function

InDebugMode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InDebugMode of Form dockSettings"
End Function
'---------------------------------------------------------------------------------------
' Procedure : btnSaveRestart_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : Apply the registry or settings.ini
'---------------------------------------------------------------------------------------
'
Private Sub btnSaveRestart_Click()
    
    Call saveOrRestart(True)

   On Error GoTo 0
   Exit Sub

btnSaveRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") " & " in procedure btnSaveRestart_Click of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : saveOrRestart
' Author    : beededea
' Date      : 24/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub saveOrRestart(ByVal restart As Boolean)

    ' variables declared
    Dim NameProcess As String
    Dim ans As Boolean
    Dim answer As VbMsgBoxResult
    Dim positionZeroFail As Boolean
    Dim positionThreeFail As Boolean
    Dim debugPoint As Integer
    Dim itis As Boolean: itis = False

    On Error GoTo saveOrRestart_Error

    If debugflg = 1 Then Debug.Print "%btnSaveRestart_Click"
   
   'initialise the dimensioned variables
    NameProcess = vbNullString
    ans = False
    answer = vbNo
    positionZeroFail = False
    positionThreeFail = False
    debugPoint = 0
    
    picBusy.Visible = True
   
    If InIDE = True Then
        If optGeneralReadRegistry.Value = True Then
            answer = MsgBox("Running in the IDE. The registry values may corrupt - be warned. Continue?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
'    Else
'        If IsUserAnAdmin() = 0 Then
'            answer = MsgBox("This program is not running as admin. Some of the settings may be strange and unwanted - be warned. Continue?", vbYesNo)
'            If answer = vbNo Then
'                Exit Sub
'            End If
'        End If
    End If
   
   
    ' if the settings.ini has been chosen as an option then the creation of it will already have occurred,
    ' so, if the temporary settings file exists then it means that the user clicked "use settings.ini file"
    ' in which case we copy it to the main settings.ini file.
    
    debugPoint = 1
    ' Steamydock exists so we shall write to the settings file those additonal items that need to be there regardless of the location of the dock data
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralReadConfig", rDGeneralReadConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralWriteConfig", rDGeneralWriteConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "RunAppInterval", rDRunAppInterval, dockSettingsFile
    'PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "DefaultDock", rDDefaultDock, dockSettingsFile
    
    debugPoint = 2

    ' writes to the new config file
    Call writeDockSettings("Software\SteamyDock\DockSettings", dockSettingsFile)
    
    ' the docksettings tool only writes the basic 'dock' configuration
    ' however, if the 'icon' settings do not exist in the 3rd config option then the actual dock will fail to show any icons
    ' (the other icon settings tool is meant to write the icon data but that tool may not yet have been run).
    
    ' in the unlikely event that this program is run before the main dock program, there is a chance that the dockSettings.ini
    ' will not have been created previously and may not contain any icon details. This next bit tests the docksettings.ini
    ' file for valid icon records.
    
    'test the first record
    sFilename = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-FileName", dockSettingsFile)
    sTitle = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-Title", dockSettingsFile)
    sCommand = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-Command", dockSettingsFile)
    If sFilename = vbNullString And sTitle = vbNullString And sCommand = vbNullString Then positionZeroFail = True

    'test the third record - it assumes all docks will have at least three elements and therfore three records
    sFilename = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-FileName", dockSettingsFile)
    sTitle = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-Title", dockSettingsFile)
    sCommand = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-Command", dockSettingsFile)
    If sFilename = vbNullString And sTitle = vbNullString And sCommand = vbNullString Then positionThreeFail = True
    
    ' the dock icon settings are empty? deanieboy
    If positionZeroFail = True And positionThreeFail = True Then
        If fFExists(dockSettingsFile) Then ' does the temporary settings.ini exist?
            ' read the registry values for each of the icons and write them to the settings.ini
            'Call readIconsWriteSettings("Software\SteamyDock\IconSettings", dockSettingsFile)
        End If
    End If

    If rDStartupRun = "1" Then
        If defaultDock = 1 Then ' steamydock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteamyDock", """" & txtAppPath.Text & "\" & "SteamyDock.exe""")
        End If
    Else
        If defaultDock = 1 Then ' steamydock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteamyDock", vbNullString)
        End If
    End If
    
    totalBusyMaximum = 20
    busyTimer.Enabled = True
    
    ' only restart if reuired
    If restart = True Then
        totalBusyMaximum = 50
        ' kill the steamydock process first
        NameProcess = dockAppPath & "\" & "SteamyDock.exe" ' .07 DAEB 01/02/2021 dockSettings.frm Modified the parameter passed to isRunning to include the full path, otherwise it does not correlate with the found processes' folder
    
        itis = IsRunning(NameProcess, vbNull) ' this is the check to see if the process is running
        ' kill the rocketdock /steamydock process first
        If itis = True Then
            ans = checkAndKill(NameProcess, False, False, False)
            If ans = True Then ' only proceed if the kill has succeeded
                
                ' restart steamydock
                If fFExists(NameProcess) Then ' .09 DAEB 07/02/2021 dockSettings.frm use the fullprocess variable without adding path again - duh!
                    ans = ShellExecute(hWnd, "Open", NameProcess, vbNullString, App.Path, 1)
                End If
            End If
        Else
            answer = MsgBox("Could not find a " & NameProcess & " process, would you like me to restart " & NameProcess & "?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
    
            ' restart Rocketdock
            If fFExists(NameProcess) Then
                Call ShellExecute(hWnd, "Open", NameProcess, vbNullString, App.Path, 1)
            End If
        End If
    End If
    
    Call repositionWindowsTaskbar(rDSide, rDSide)

   On Error GoTo 0
   Exit Sub

saveOrRestart_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveOrRestart of Form dockSettings"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error
   If debugflg = 1 Then Debug.Print "%btnClose_Click"

    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnDefaults_Click
' Author    : beededea
' Date      : 02/03/2020
' Purpose   : The registry is written first and then the settings file is recreated afterwards
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaults_Click()

        
    ' variables declared
    Dim NameProcess As String
    Dim ans As Boolean
    Dim answer As VbMsgBoxResult

   On Error GoTo btnDefaults_Click_Error
   If debugflg = 1 Then Debug.Print "%btnDefaults_Click"

   'initialise the dimensioned variables
    NameProcess = vbNullString
    ans = False
    answer = vbNo
    
    If InIDE = True Then
        answer = MsgBox("Running in the IDE. The registry values may corrupt - be warned. Continue?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    Else
        If IsUserAnAdmin() = 0 Then
            answer = MsgBox("This program is not running as admin. Some of the settings may be strange and unwanted - be warned. Continue?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If

    answer = MsgBox("Are you sure you want to rest Rocketdock to its default settings? Note: this will not lose your icons.?", vbYesNo)
    If answer = vbNo Then
        Exit Sub
    End If
 
    ' kill the rocketdock process
    
'    If defaultDock = 0 Then
'        NameProcess = "RocketDock.exe"
'    Else
'        NameProcess = "steamyDock.exe"
'    End If
'    ans = checkAndKill(NameProcess, False)
    
'    If defaultDock = 0 Then
'        rDVersion = "1.3.5"
'    Else
        rDVersion = App.Major & "." & App.Minor & "." & App.Revision
'    End If
    
    rDCustomIconFolder = vbNullString
    
'    If defaultDock = 0 Then ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'        rDHotKeyToggle = "Control+Alt+R"
'    Else
        rDHotKeyToggle = "F11"
'    End If
    cmbHidingKey.Text = rDHotKeyToggle
        
    ' removed
    'cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
'    If defaultDock = 1 Then
'        cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'    Else
'        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'    End If
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
   
    
    rDtheme = "CrystalXP.net"
    rDWallpaper = "none selected"
    rDWallpaperStyle = "Centre"
    rDAutomaticWallpaperChange = "0"
    rDWallpaperTimerIntervalIndex = "5"
    rDWallpaperTimerInterval = "60"
    
    rDMoveWinTaskbar = "1"

    cmbStyleTheme.Text = rDtheme
    
    rDThemeOpacity = "100"
    sliStyleOpacity.Value = Val(rDThemeOpacity)
    
    rDIconOpacity = "100"
    sliIconsOpacity.Value = Val(rDIconOpacity)
    
    rDFontSize = "-8"
    rDFontFlags = "0"
    rDFontName = "Centurion Light SF"
    rDFontColor = "65535"
    rDFontCharSet = "1"
    
    'lblPreviewFont.ForeColor = Convert_Dec2RGB(rDFontColor)  ' converts the stored decimal value to RGB

    lblPreviewFont.FontName = rDFontName
    lblPreviewFont.FontSize = Abs(rDFontSize)
    lblPreviewFont.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewFont.ForeColor = rDFontColor
    
    lblPreviewTop.FontName = rDFontName
    lblPreviewTop.FontSize = Abs(rDFontSize)
    lblPreviewTop.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewBottom.FontName = rDFontName
    lblPreviewBottom.FontSize = Abs(rDFontSize)
    lblPreviewBottom.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewLeft.FontName = rDFontName
    lblPreviewLeft.FontSize = Abs(rDFontSize)
    lblPreviewLeft.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewRight.FontName = rDFontName
    lblPreviewRight.FontSize = Abs(rDFontSize)
    lblPreviewRight.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewFontShadow.FontName = rDFontName
    lblPreviewFontShadow.FontSize = Abs(rDFontSize)
    lblPreviewFontShadow.FontBold = rDFontFlags
    'lblPreviewFontShadow.FontItalic = suppliedStyle
    lblPreviewFontShadow.ForeColor = rDFontShadowColor
    
    lblPreviewFontShadow2.FontName = rDFontName
    lblPreviewFontShadow2.FontSize = Abs(rDFontSize)
    lblPreviewFontShadow2.FontBold = rDFontFlags
    'lblPreviewFontShadow2.FontItalic = suppliedStyle
    lblPreviewFontShadow2.ForeColor = rDFontShadowColor
    
    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Abs(rDFontSize) & "pt"
           
    rDFontOutlineColor = "255"
    lblStyleOutlineColourDesc.Caption = "Outline Colour: " & Convert_Dec2RGB(rDFontOutlineColor)
    lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor
    
    rDFontOutlineOpacity = "0"
    sliStyleOutlineOpacity.Value = Val(rDFontOutlineOpacity)
    
    rDFontShadowColor = "12632256"
    lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & Convert_Dec2RGB(rDFontShadowColor)
    lblStyleFontOutlineTest.ForeColor = rDFontShadowColor
    
    rDFontShadowOpacity = "30"
    sliStyleShadowOpacity.Value = Val(rDFontShadowOpacity)
    
    sDFontOpacity = "100"
    sliStyleFontOpacity.Value = Val(sDFontOpacity)
    
    rDIconMin = "16"
    sliIconsSize.Value = Val(rDIconMin)
     
    sliIconsZoom.Value = Val(rdIconMax) - 17
    
    Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animatio
    
    rDZoomWidth = "4"
    sliIconsZoomWidth.Value = Val(rDZoomWidth)
    
    rDZoomTicks = "199"
    sliIconsDuration.Value = Val(rDZoomTicks)
    
    rDAutoHideDuration = "186"
    sliAutoHideDuration.Value = Val(rDAutoHideDuration)
    
    rDAnimationInterval = "10"
    sliAnimationInterval.Value = Val(rDAnimationInterval)
    
    rDSkinSize = "118"
    sliStyleThemeSize.Value = Val(rDSkinSize)
    
    sDSplashStatus = "1"
    chkSplashStatus.Value = Val(sDSplashStatus)
    
    sDShowIconSettings = "1"
    chkShowIconSettings.Value = Val(sDShowIconSettings) ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock
    
    rDAutoHideDelay = "174"
    sliBehaviourAutoHideDelay.Value = Val(rDAutoHideDelay)
    
    rDPopupDelay = "68"
    sliBehaviourPopUpDelay.Value = Val(rDPopupDelay)
    
    sDContinuousHide = "10" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    sliContinuousHide.Value = sDContinuousHide
    
    sDBounceZone = "75"
    ' sDBounceZone
    
    rDIconQuality = "2"
    cmbIconsQuality.ListIndex = Val(rDIconQuality)
    
    rDLangID = "1033"
    
    rDHideLabels = "0"
    chkStyleDisable.Value = Val(rDHideLabels)
    
   ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
    sDShowLblBacks = "0"
    chkLabelBackgrounds.Value = Val(sDShowLblBacks)

    
    rDZoomOpaque = "1"
    chkIconsZoomOpaque.Value = Val(rDZoomOpaque)
    
    rDLockIcons = "1"
    chkLockIcons.Value = Val(rDLockIcons)
    
    ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    rDRetainIcons = "0"
    chkRetainIcons.Value = Val(rDRetainIcons)
    
    rDAutoHide = "1"
    chkAutoHide.Value = Val(rDAutoHide)
' 26/10/2020 docksettings .05 DAEB  added a manual click to the autohide toggle checkbox
' a checkbox value assignment does not trigger a checkbox click for this checkbox (in a frame) as normally occurs and there is no equivalent 'change event' for a checkbox
' so to force it to trigger we need a call to the click event
    Call chkAutoHide_Click
    
    rDManageWindows = "0"
    chkGenMin.Value = Val(rDManageWindows)
    
    rDDisableMinAnimation = "1"
    chkGenDisableAnim.Value = Val(rDDisableMinAnimation)
    
    rDShowRunning = "1"
    chkShowRunning.Value = Val(rDShowRunning)
    
    rDOpenRunning = "1"
    chkOpenRunning.Value = Val(rDOpenRunning)
    
    rDHoverFX = "1"
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)
    
    rDzOrderMode = "0"
    cmbPositionLayering.ListIndex = Val(rDzOrderMode)
    
    rDMouseActivate = "1"
    chkBehaviourMouseActivate.Value = Val(rDMouseActivate)
    
    rDIconActivationFX = "2"
    cmbIconActivationFX.ListIndex = Val(rDIconActivationFX)
    
    rDSoundSelection = "0"
    cmbBehaviourSoundSelection.ListIndex = Val(rDSoundSelection)
    
    sDAutoHideType = "0"
    cmbAutoHideType.ListIndex = Val(sDAutoHideType)
    
    rDMonitor = "0" ' ie. monitor 1
    cmbPositionMonitor.ListIndex = Val(rDMonitor)
    
    rDSide = "1"
    cmbPositionScreen.ListIndex = Val(rDSide)
    
    rDOffset = "0"
    sliPositionCentre.Value = Val(rDOffset)
    
    rDvOffset = "0"
    sliPositionEdgeOffset.Value = Val(rDvOffset)
    
    rDOptionsTabIndex = "4"
    rdIconMax = "128" ' 128 rocketdock
    
    ' The following has been commented out as the reversion to defaults should happen using the temporary settings file, not the registry

'    If rocketDockInstalled = True Then ' if RD hasn't been installed then the registry nor the settings.ini file will exist
'
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMax", "128")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LoadError", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Version", "1.3.5")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "CustomIconFolder", "?E:\\dean\\steampunk theme\\icons")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HotKey-Toggle", "Control+Alt+R")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Theme", "CrystalXP.net")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ThemeOpacity", "100")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconOpacity", "100")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontSize", "-8")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontFlags", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontName", "Times New Roman")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontColor", "65535")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontCharSet", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineColor", "255")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineOpacity", "9")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowColor", "12632256")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowOpacity", "30")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMin", "16")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomWidth", "4")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomTicks", "199")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideTicks", "186")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideDelay", "174")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "PopupDelay", "68")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconQuality", "2")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LangID", "1033")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HideLabels", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomOpaque", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHide", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ManageWindows", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "DisableMinAnimation", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ShowRunning", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OpenRunning", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HoverFX", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "zOrderMode", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "MouseActivate", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconActivationFX", "2")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Monitor", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Offset", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "vOffset", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OptionsTabIndex", "4")
'    End If
    
    'regardless of the method used, registry, settings or the new 3rd option, they all use the temporary settings file.
    If fFExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
        Call writeDockSettings("Software\RocketDock", tmpSettingsFile)
        ' if it exists, read the registry values for each of the icons and write them to the settings.ini
        Call readIconsWriteSettings("Software\RocketDock\Icons", tmpSettingsFile)
        
    End If
    
    ' writes directly to the new config file without any intervening temporary file
    If optGeneralReadConfig.Value = True Then
        Call writeDockSettings("Software\SteamyDock\DockSettings", dockSettingsFile)
    End If

    'NOTE: the settings are NOT written to the registry until the apply button is pressed.
    
    ' if the rocketdock process has died then
'    If ans = True Then
'        ' restart Rocketdock
'        Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
'    Else
'
'        answer = MsgBox("Could not find a " & NameProcess & " process, would you like me to restart " & NameProcess & "?", vbYesNo)
'        If answer = vbNo Then
'            Exit Sub
'        End If
'
'        ' restart Rocketdock
'        Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
'    End If

   On Error GoTo 0
   Exit Sub

btnDefaults_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaults_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGeneralRdFolder_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   : unused disabled
'---------------------------------------------------------------------------------------
'
Private Sub btnGeneralRdFolder_Click()

        
    ' variables declared
    Dim getFolder As String
    Dim dialogInitDir As String
   
    On Error GoTo btnGeneralRdFolder_Click_Error
    If debugflg = 1 Then Debug.Print "%btnGeneralRdFolder_Click"
    
   'initialise the dimensioned variables
    getFolder = vbNullString
    dialogInitDir = vbNullString
    
    If txtAppPath.Text <> vbNullString Then
        If fDirExists(txtAppPath.Text) Then
            dialogInitDir = txtAppPath.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder
    'getFolder = ChooseDir_Click ' old method to show the dialog box to select a folder
    If getFolder <> vbNullString Then txtAppPath.Text = getFolder

   On Error GoTo 0
   Exit Sub

btnGeneralRdFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGeneralRdFolder_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStyleFont_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   : select the dock's font
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleFont_Click()
        
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedFontSize As Integer
    
    Dim suppliedStyle As Boolean
    Dim suppliedColour As Variant
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    Dim fontSelected As Boolean
    
    On Error GoTo btnStyleFont_Click_Error
    If debugflg = 1 Then Debug.Print "%btnStyleFont_Click"
   
   'initialise the dimensioned variables
    suppliedFont = vbNullString
    suppliedSize = 0
    suppliedWeight = 0
    suppliedBold = False
    suppliedFontSize = 0
    
    suppliedStyle = False
    'suppliedColour =
    suppliedItalics = False
    suppliedUnderline = False
    fontSelected = False
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    
    displayFontSelector rDFontName, suppliedFontSize, suppliedWeight, suppliedStyle, rDFontColor, suppliedItalics, suppliedUnderline, fontSelected
    If fontSelected = False Then Exit Sub
    
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    
    
   On Error GoTo 0
   Exit Sub

btnStyleFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleFont_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : preFontInformation
' Author    : beededea
' Date      : 17/05/2020
' Purpose   : the fontsize used by Rocketdock does not equate to the pt size in the font selector
'             so we have to determine the fontsize to be displayed by the font selector
'---------------------------------------------------------------------------------------
'
Private Sub preFontInformation(suppliedFontSize As Integer, suppliedBold As Boolean, suppliedItalics As Boolean, suppliedUnderline As Boolean, suppliedWeight As Integer)

   On Error GoTo preFontInformation_Error

    If rDFontSize = "-8" Then suppliedFontSize = 6
    If rDFontSize = "-11" Then suppliedFontSize = 8
    If rDFontSize = "-12" Then suppliedFontSize = 9
    If rDFontSize = "-13" Then suppliedFontSize = 10
    If rDFontSize = "-15" Then suppliedFontSize = 11
    If rDFontSize = "-16" Then suppliedFontSize = 12
    If rDFontSize = "-19" Then suppliedFontSize = 14
    If rDFontSize = "-21" Then suppliedFontSize = 16
    If rDFontSize = "-24" Then suppliedFontSize = 18
    If rDFontSize = "-27" Then suppliedFontSize = 20
    If rDFontSize = "-29" Then suppliedFontSize = 22
    
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False
    'suppliedWeight = False
    
    If rDFontFlags = 1 Or rDFontFlags = 3 Or rDFontFlags = 7 Or rDFontFlags = 11 Or rDFontFlags = 15 Then suppliedBold = True
    If rDFontFlags = 2 Or rDFontFlags = 3 Or rDFontFlags = 6 Or rDFontFlags = 7 Or rDFontFlags = 10 Or rDFontFlags = 11 Or rDFontFlags = 13 Or rDFontFlags = 14 Or rDFontFlags = 15 Then suppliedItalics = True
    If rDFontFlags = 6 Or rDFontFlags = 14 Then suppliedUnderline = True


   On Error GoTo 0
   Exit Sub

preFontInformation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure preFontInformation of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontInformation
' Author    : beededea
' Date      : 17/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayFontInformation(suppliedFontSize As Integer, suppliedBold As Boolean, suppliedItalics As Boolean, suppliedUnderline As Boolean, suppliedWeight As Integer)
    
    ' the fontsize used by Rocketdock does not equate to the pt size in the font selector
    ' so we have to calculate the rDFontSize that will be written to rocketdock registry/settings.
    
   On Error GoTo displayFontInformation_Error

    If suppliedFontSize = 6 Then rDFontSize = "-8"
    If suppliedFontSize = 8 Then rDFontSize = "-11"
    If suppliedFontSize = 9 Then rDFontSize = "-12"
    If suppliedFontSize = 10 Then rDFontSize = "-13"
    If suppliedFontSize = 11 Then rDFontSize = "-15"
    If suppliedFontSize = 12 Then rDFontSize = "-16"
    If suppliedFontSize = 14 Then rDFontSize = "-19"
    If suppliedFontSize = 16 Then rDFontSize = "-21"
    If suppliedFontSize = 18 Then rDFontSize = "-24"
    If suppliedFontSize = 20 Then rDFontSize = "-27"
    If suppliedFontSize = 22 Then rDFontSize = "-29"
    
    lblPreviewFont.FontName = rDFontName
    lblPreviewFont.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewTop.FontName = rDFontName
    lblPreviewTop.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewBottom.FontName = rDFontName
    lblPreviewBottom.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewLeft.FontName = rDFontName
    lblPreviewLeft.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewRight.FontName = rDFontName
    lblPreviewRight.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewFontShadow.FontName = rDFontName
    lblPreviewFontShadow.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewFontShadow2.FontName = rDFontName
    lblPreviewFontShadow2.FontSize = Abs(suppliedFontSize) + 4
    
    If suppliedWeight > 400 Then
        suppliedBold = True
    Else
        suppliedBold = False
    End If
    
    lblPreviewFont.FontBold = suppliedBold
    lblPreviewFont.FontItalic = suppliedItalics
    lblPreviewFont.ForeColor = rDFontColor
    lblPreviewFont.FontUnderline = suppliedUnderline
    
    lblPreviewTop.FontBold = suppliedBold
    lblPreviewTop.FontItalic = suppliedItalics
    lblPreviewTop.ForeColor = rDFontOutlineColor
    lblPreviewTop.FontUnderline = suppliedUnderline
    
    lblPreviewBottom.FontBold = suppliedBold
    lblPreviewBottom.FontItalic = suppliedItalics
    lblPreviewBottom.ForeColor = rDFontOutlineColor
    lblPreviewBottom.FontUnderline = suppliedUnderline
    
    lblPreviewLeft.FontBold = suppliedBold
    lblPreviewLeft.FontItalic = suppliedItalics
    lblPreviewLeft.ForeColor = rDFontOutlineColor
    lblPreviewLeft.FontUnderline = suppliedUnderline
    
    lblPreviewRight.FontBold = suppliedBold
    lblPreviewRight.FontItalic = suppliedItalics
    lblPreviewRight.ForeColor = rDFontOutlineColor
    lblPreviewRight.FontUnderline = suppliedUnderline
    
    lblPreviewFontShadow.FontBold = suppliedBold
    lblPreviewFontShadow.FontItalic = suppliedItalics
    lblPreviewFontShadow.ForeColor = rDFontShadowColor
    lblPreviewFontShadow.FontUnderline = suppliedUnderline
    
    lblPreviewFontShadow2.FontBold = suppliedBold
    lblPreviewFontShadow2.FontItalic = suppliedItalics
    lblPreviewFontShadow2.ForeColor = rDFontShadowColor
    lblPreviewFontShadow2.FontUnderline = suppliedUnderline
    
    'lblPreviewFontShadow.Visible = False
    
    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Abs(suppliedFontSize) & "pt"
    If suppliedBold = True Then lblStyleFontName.Caption = lblStyleFontName.Caption & " Bold"
    If suppliedItalics = True Then lblStyleFontName.Caption = lblStyleFontName.Caption & " Italic"

    ' now change the rocketdock vars
    
    ' 0 - no qualifiers or alterations
    ' 1 - bold
    ' 2 - light italics
    ' 3 - bold italics
    ' 4 - strikeout & light ' unsupported
    ' 6 - underline and italics
    ' 7 - bold, italics & underline
    ' 10 - strikeout & italics ' unsupported
    ' 11 - bold, italics & strikeout  ' unsupported
    ' 13 - strikeout & italics        ' unsupported
    ' 14 - underline, strikeout and italics ' unsupported
    ' 15 - bold, underline, strikeout and italics ' unsupported
        
    lblPreviewFont.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewFont.top = (735 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewTop.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewTop.top = (715 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewBottom.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewBottom.top = (755 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewLeft.Left = (5320 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewLeft.top = (735 / 2) - (lblPreviewFont.Height / 2)
        
    lblPreviewRight.Left = (5370 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewRight.top = (735 / 2) - (lblPreviewFont.Height / 2)
        
    lblPreviewFontShadow.Left = (5440 / 2) - (lblPreviewFontShadow.Width / 2)
    lblPreviewFontShadow.top = (825 / 2) - (lblPreviewFontShadow.Height / 2)
        
    lblPreviewFontShadow2.Left = (5450 / 2) - (lblPreviewFontShadow2.Width / 2)
    lblPreviewFontShadow2.top = (835 / 2) - (lblPreviewFontShadow2.Height / 2)
        
    rDFontFlags = 0
    If suppliedBold = True Then rDFontFlags = 1
    If suppliedItalics = True Then rDFontFlags = 2
    If suppliedItalics = True And suppliedBold = True Then rDFontFlags = 3
    If suppliedUnderline = True And suppliedItalics = True Then rDFontFlags = 6
    If suppliedUnderline = True And suppliedItalics = True And suppliedBold = True Then rDFontFlags = 7

   On Error GoTo 0
   Exit Sub

displayFontInformation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontInformation of Form dockSettings"
  End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStyleShadow_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   : determine the shadow colour
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleShadow_Click()
        
    ' variables declared
    Dim colourResult As Long
    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
   
    'initialise the dimensioned variables
     colourResult = 0
     suppliedFontSize = 0
     suppliedWeight = 0
     suppliedBold = False
     suppliedItalics = False
     suppliedUnderline = False
   
   On Error GoTo btnStyleShadow_Click_Error
   If debugflg = 1 Then Debug.Print "%btnStyleShadow_Click"

    colourResult = ShowColorDialog(Me.hWnd, True, rDFontShadowColor)

    If colourResult <> -1 And colourResult <> 0 Then
        rDFontShadowColor = colourResult
        
        lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & Convert_Dec2RGB(rDFontShadowColor)
        lblStyleFontFontShadowTest.ForeColor = rDFontShadowColor
        
        'rDFontShadowOpacity = str(btnStyleShadow.value)

    End If
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
       

    On Error GoTo 0
    Exit Sub

btnStyleShadow_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleShadow_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnStyleOutline_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleOutline_Click()

       
    ' variables declared
    Dim colourResult As Long
    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    
    'initialise the dimensioned variables
     colourResult = 0
     suppliedFontSize = 0
     suppliedWeight = 0
     suppliedBold = False
     suppliedItalics = False
     suppliedUnderline = False
    
    On Error GoTo btnStyleOutline_Click_Error
   If debugflg = 1 Then Debug.Print "%btnStyleOutline_Click"
    
    ' this will take 255, VBRed,  16711680
    colourResult = ShowColorDialog(Me.hWnd, True, rDFontOutlineColor)
    
    If colourResult <> -1 And colourResult <> 0 Then
        rDFontOutlineColor = (colourResult)
        lblStyleOutlineColourDesc.Caption = "Outline Colour: " & Convert_Dec2RGB(rDFontOutlineColor)
        lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor
    End If
   
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
   
   On Error GoTo 0
   Exit Sub

btnStyleOutline_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleOutline_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkAutoHide_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutoHide_Click()
   On Error GoTo chkAutoHide_Click_Error
   If debugflg = 1 Then Debug.Print "%chkAutoHide_Click"

    If chkAutoHide.Value = 1 Then
        chkAutoHide.Caption = "Autohide Enabled"
        sliAutoHideDuration.Enabled = True
  
        lblBehaviourLabel(2).Enabled = True

        lblBehaviourLabel(8).Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblBehaviourLabel(4).Enabled = True
        lblBehaviourLabel(10).Enabled = True
        sliBehaviourAutoHideDelay.Enabled = True
        lblAutoHideDelayMsHigh.Enabled = True
        lblAutoHideDelayMsCurrent.Enabled = True
        
        lblBehaviourLabel(3).Enabled = True
        lblBehaviourLabel(9).Enabled = True
        lblAutoRevealDurationMsHigh.Enabled = True
        sliBehaviourPopUpDelay.Enabled = True
        
        lblBehaviourPopUpDelayMsCurrrent.Enabled = True
        
        cmbAutoHideType.Enabled = True
        
    Else
        chkAutoHide.Caption = "Autohide Disabled"
        sliAutoHideDuration.Enabled = False

        lblBehaviourLabel(2).Enabled = False

        lblBehaviourLabel(8).Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblBehaviourLabel(4).Enabled = False
        lblBehaviourLabel(10).Enabled = False
        sliBehaviourAutoHideDelay.Enabled = False
        lblAutoHideDelayMsHigh.Enabled = False
        lblAutoHideDelayMsCurrent.Enabled = False
        
                
        lblBehaviourLabel(3).Enabled = False
        lblBehaviourLabel(9).Enabled = False
        sliBehaviourPopUpDelay.Enabled = False
        lblAutoRevealDurationMsHigh.Enabled = False
        lblBehaviourPopUpDelayMsCurrrent.Enabled = False
        
        cmbAutoHideType.Enabled = False
    
    End If
    
    rDAutoHide = chkAutoHide.Value
    

   On Error GoTo 0
   Exit Sub

chkAutoHide_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAutoHide_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkBehaviourMouseActivate_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkBehaviourMouseActivate_Click()
   On Error GoTo chkBehaviourMouseActivate_Click_Error
   If debugflg = 1 Then Debug.Print "%chkBehaviourMouseActivate_Click"

    rDMouseActivate = chkBehaviourMouseActivate.Value

   On Error GoTo 0
   Exit Sub

chkBehaviourMouseActivate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkBehaviourMouseActivate_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenDisableAnim_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenDisableAnim_Click()

   On Error GoTo chkGenDisableAnim_Click_Error
   If debugflg = 1 Then Debug.Print "%chkGenDisableAnim_Click"

   rDDisableMinAnimation = chkGenDisableAnim.Value
   
   rDMouseActivate = chkGenDisableAnim.Value

   On Error GoTo 0
   Exit Sub

chkGenDisableAnim_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenDisableAnim_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkLockIcons_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkLockIcons_Click()

   On Error GoTo chkLockIcons_Click_Error
   If debugflg = 1 Then Debug.Print "%chkLockIcons_Click"

    rDLockIcons = chkLockIcons.Value

   On Error GoTo 0
   Exit Sub

chkLockIcons_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkLockIcons_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenMin_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenMin_Click()
   On Error GoTo chkGenMin_Click_Error
   If debugflg = 1 Then Debug.Print "%chkGenMin_Click"

    If chkGenMin.Value = 0 Then
        chkGenDisableAnim.Enabled = False
    Else
        chkGenDisableAnim.Enabled = True
    End If
    
    rDManageWindows = chkGenMin.Value

   On Error GoTo 0
   Exit Sub

chkGenMin_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenMin_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkOpenRunning_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkOpenRunning_Click()
   On Error GoTo chkOpenRunning_Click_Error
   If debugflg = 1 Then Debug.Print "%chkOpenRunning_Click"

    rDOpenRunning = chkOpenRunning.Value

   On Error GoTo 0
   Exit Sub

chkOpenRunning_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkOpenRunning_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkShowRunning_Click
' Author    : beededea
' Date      : 11/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkShowRunning_Click()
   On Error GoTo chkShowRunning_Click_Error
   If debugflg = 1 Then Debug.Print "%chkShowRunning_Click"

    rDShowRunning = chkShowRunning.Value
    
    If chkShowRunning.Value = 0 Then
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else

        If optGeneralReadConfig.Value = True Then ' steamydock
'            lblGenLabel(0).Enabled = True
'            lblGenLabel(1).Enabled = True
            sliRunAppInterval.Enabled = True
            lblGenLabel(2).Enabled = True
            lblGenRunAppIntervalCur.Enabled = True
        End If
    End If

   On Error GoTo 0
   Exit Sub

chkShowRunning_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowRunning_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkStartupRun_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkStartupRun_Click()

   On Error GoTo chkStartupRun_Click_Error
   If debugflg = 1 Then Debug.Print "%chkStartupRun_Click"

    rDStartupRun = chkStartupRun.Value

   On Error GoTo 0
   Exit Sub

chkStartupRun_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkStartupRun_Click of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkIconsZoomOpaque_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkIconsZoomOpaque_Click()

   On Error GoTo chkIconsZoomOpaque_Click_Error
   If debugflg = 1 Then Debug.Print "%chkIconsZoomOpaque_Click"

    rDZoomOpaque = chkIconsZoomOpaque.Value

   On Error GoTo 0
   Exit Sub

chkIconsZoomOpaque_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIconsZoomOpaque_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkStyleDisable_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkStyleDisable_Click()

   On Error GoTo chkStyleDisable_Click_Error
   If debugflg = 1 Then Debug.Print "%chkStyleDisable_Click"

   rDHideLabels = chkStyleDisable.Value
   
    If chkStyleDisable.Value = 1 Then
        chkLabelBackgrounds.Enabled = False ' .01 docksettings DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
                
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkLabelBackgrounds.Width = 192 ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblChkLabelBackgrounds.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblChkLabelBackgrounds.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        btnStyleFont.Enabled = False
        lblStyleFontName.Enabled = False
        btnStyleShadow.Enabled = False
        lblStyleFontFontShadowColor.Enabled = False
        lblStyleFontFontShadowTest.Enabled = False
        btnStyleOutline.Enabled = False
        lblStyleOutlineColourDesc.Enabled = False
        lblStyleFontOutlineTest.Enabled = False
        lblStyleLabel(3).Enabled = False
        lblStyleLabel(4).Enabled = False
        lblStyleLabel(5).Enabled = False
        lblStyleLabel(8).Enabled = False
        lblStyleLabel(9).Enabled = False
        lblStyleLabel(10).Enabled = False
        sliStyleShadowOpacity.Enabled = False
        lblStyleShadowOpacityCurrent.Enabled = False
        sliStyleOutlineOpacity.Enabled = False
        lblStyleOutlineOpacityCurrent.Enabled = False
        
        sliStyleFontOpacity.Enabled = False
        lblStyleFontOpacityCurrent.Enabled = False
        
    Else
        chkLabelBackgrounds.Enabled = True  ' .01 docksettings DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
        'lblChkLabelBackgrounds.Enabled = True ' .01
        
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkLabelBackgrounds.Width = 2490 ' set the width to show the full check box and its intrinsic label
        lblChkLabelBackgrounds.Visible = False ' make the associated duplicate label hidden
        lblChkLabelBackgrounds.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        btnStyleFont.Enabled = True
        lblStyleFontName.Enabled = True
        btnStyleShadow.Enabled = True
        lblStyleFontFontShadowColor.Enabled = True
        lblStyleFontFontShadowTest.Enabled = True
        btnStyleOutline.Enabled = True
        lblStyleOutlineColourDesc.Enabled = True
        lblStyleFontOutlineTest.Enabled = True
        sliStyleShadowOpacity.Enabled = True
        lblStyleShadowOpacityCurrent.Enabled = True
        lblStyleLabel(3).Enabled = True
        lblStyleLabel(4).Enabled = True
        lblStyleLabel(5).Enabled = True
        lblStyleLabel(8).Enabled = True
        lblStyleLabel(9).Enabled = True
        lblStyleLabel(10).Enabled = True
        sliStyleOutlineOpacity.Enabled = True
        lblStyleOutlineOpacityCurrent.Enabled = True
                
        sliStyleFontOpacity.Enabled = True
        lblStyleFontOpacityCurrent.Enabled = True
        
    End If
   

   On Error GoTo 0
   Exit Sub

chkStyleDisable_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkStyleDisable_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbIconActivationFX_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbIconActivationFX_Click()

   On Error GoTo cmbIconActivationFX_Click_Error
   If debugflg = 1 Then Debug.Print "%cmbIconActivationFX_Click"

    rDIconActivationFX = cmbIconActivationFX.ListIndex

   On Error GoTo 0
   Exit Sub

cmbIconActivationFX_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbIconActivationFX_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbBehaviourSoundSelection_Click
' Author    : beededea
' Date      : 17/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbBehaviourSoundSelection_Click()
    On Error GoTo cmbBehaviourSoundSelection_Click_Error

    rDSoundSelection = cmbBehaviourSoundSelection.ListIndex

    On Error GoTo 0
    Exit Sub

cmbBehaviourSoundSelection_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbBehaviourSoundSelection_Click of Form dockSettings"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbIconsHoverFX_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbIconsHoverFX_Click()

   On Error GoTo cmbIconsHoverFX_Change_Error
   If debugflg = 1 Then Debug.Print "%cmbIconsHoverFX_Change"

    rDHoverFX = cmbIconsHoverFX.ListIndex
    
    rDHoverFX = "1"  'DEAN needs to be removed later
    
    'none
    'bubble
    'plateau
    'flat
    'bumpy
    
    'Call setMinimumHoverFX    ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation


   On Error GoTo 0
   Exit Sub

cmbIconsHoverFX_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbIconsHoverFX_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setMinimumHoverFX
' Author    : beededea
' Date      : 29/04/2021
' Purpose   : Set the large icon minimum size to 85 pixels when using the bumpy animation
'---------------------------------------------------------------------------------------
'
Private Sub setMinimumHoverFX()
    On Error GoTo setMinimumHoverFX_Error

    If Val(rDHoverFX) = 4 And sliIconsZoom.Value <= 85 Then
        sliIconsZoom.Value = 85
        
        If chkToggleDialogs.Value = 0 Then
            sliIconsZoom.ToolTipText = "The maximum size after a zoom can be no smaller than 85 pixels when Zoom:Bumpy is chosen"
        Else
            sliIconsZoom.ToolTipText = vbNullString
        End If
    Else
        
        If chkToggleDialogs.Value = 0 Then
            sliIconsZoom.ToolTipText = "The maximum size after a zoom"
        Else
            sliIconsZoom.ToolTipText = vbNullString
        End If
    End If

    On Error GoTo 0
    Exit Sub

setMinimumHoverFX_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMinimumHoverFX of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbIconsQuality_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbIconsQuality_Click()

   On Error GoTo cmbIconsQuality_Change_Error
   If debugflg = 1 Then Debug.Print "%cmbIconsQuality_Change"
    
    rDIconQuality = cmbIconsQuality.ListIndex

   On Error GoTo 0
   Exit Sub

cmbIconsQuality_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbIconsQuality_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionLayering_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionLayering_Click()


   On Error GoTo cmbPositionLayering_Click_Error
   If debugflg = 1 Then Debug.Print "%cmbPositionLayering_Click"

   rDzOrderMode = cmbPositionLayering.ListIndex


   On Error GoTo 0
   Exit Sub

cmbPositionLayering_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionLayering_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionMonitor_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : This is called twice during startup by GetMonitorCount adjustControls, only one really required
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionMonitor_Click()

   On Error GoTo cmbPositionMonitor_Click_Error
   If debugflg = 1 Then Debug.Print "%cmbPositionMonitor_Click"
    
    If startupFlg = True Then '
        ' don't do this on the startup run only when actually clicked upon
        Exit Sub
    Else
        rDMonitor = cmbPositionMonitor.ListIndex
    End If

   On Error GoTo 0
   Exit Sub

cmbPositionMonitor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionMonitor_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionScreen_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : This routine will place taskbar where the dock isn't to avoid overlap
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionScreen_Click()
'    Dim taskbarPosition As Integer: taskbarPosition = 0
'    Dim triggerTaskbarChange As Boolean: triggerTaskbarChange = False
'    Dim rmessage As String: rmessage = ""
'    Dim answer As VbMsgBoxResult: answer = vbNo
   
    On Error GoTo cmbPositionScreen_Click_Error
    If debugflg = 1 Then Debug.Print "%cmbPositionScreen_Click"
    
    If startupFlg = True Then '
        ' don't do this on the first startup run
        Exit Sub
    End If
    
    rDSide = CStr(cmbPositionScreen.ListIndex)
    
    ' steamydock left and right positions unsupported
    If rDSide = "2" Or rDSide = "3" Then
        rDSide = "1"
        cmbPositionScreen.ListIndex = 1
    End If

   
   On Error GoTo 0
   Exit Sub

cmbPositionScreen_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionScreen_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmbStyleTheme_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : if a theme is selected from the dropdown list then make it the default
'---------------------------------------------------------------------------------------
'
Private Sub cmbStyleTheme_Click()
    Dim themePic As String

    On Error GoTo cmbStyleTheme_Change_Error
    If debugflg = 1 Then Debug.Print "%cmbStyleTheme_Change"
    
    rDtheme = cmbStyleTheme.List(cmbStyleTheme.ListIndex)
    
    ' .09 DAEB 01/02/2021 docksettings Make the sample image functionality disabled for rocketdock
    If defaultDock = 1 Then
        themePic = sdAppPath & "\skins\" & rDtheme & "\sample.jpg"
        
        If fFExists(themePic) Then
            imgThemeSample.Picture = LoadPicture(sdAppPath & "\skins\" & rDtheme & "\sample.jpg")
        End If
    End If
    
    On Error GoTo 0
    Exit Sub

cmbStyleTheme_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbStyleTheme_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWallpaper_Click
' Author    : beededea
' Date      : 01/03/2025
' Purpose   : if a wallpaper image is selected from the dropdown list then make it the default
'---------------------------------------------------------------------------------------
'
Private Sub cmbWallpaper_Click()
    Dim wallpaperPic As String: wallpaperPic = vbNullString

    On Error GoTo cmbWallpaper_Change_Error
    If debugflg = 1 Then Debug.Print "%cmbWallpaper_Change"
    
    If startupFlg = True Then Exit Sub
    
    rDWallpaper = cmbWallpaper.List(cmbWallpaper.ListIndex)
    
    ' disable the apply button if no wallpaper choice
    If rDWallpaper <> "none selected" Then
        btnApplyWallpaper.Enabled = True
    Else
        btnApplyWallpaper.Enabled = False
    End If
    
    wallpaperPic = sdAppPath & "\wallpapers\" & rDWallpaper
    
    If fFExists(wallpaperPic) Then
        imgWallpaperPreview.Picture = LoadPicture(sdAppPath & "\wallpapers\" & rDWallpaper)
    End If

    
    On Error GoTo 0
    Exit Sub

cmbWallpaper_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWallpaper_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDefaultDock_Click
' Author    : beededea
' Date      : 13/05/2020
' Purpose   : certain options are disabled when selecting Steamydock (and vice versa)
'---------------------------------------------------------------------------------------
'
Private Sub cmbDefaultDock_Click()
   On Error GoTo cmbDefaultDock_Click_Error
   If debugflg = 1 Then Debug.Print "%cmbDefaultDock_Click"

'    If cmbDefaultDock.List(cmbDefaultDock.ListIndex) = "RocketDock" Then
'        ' check where rocketdock is installed
'        Call checkRocketdockInstallation
'        defaultDock = 0 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
'
'        ' .17 DAEB 07/09/2022 docksettings the dock folder location now changes as it is switched between Rocketdock and Steamy Dock
'        dockAppPath = rdAppPath
'        txtAppPath.Text = rdAppPath
'
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'            optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'        Else
'            optGeneralReadRegistry.Value = True
'            'optGeneralWriteRegistry.Value = True
'        End If
'
'        rDDefaultDock = "rocketdock"
'
'        ' re-enable all the controls that Rocketdock supports
'
'        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenMin.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenMin.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenMin.Visible = False ' make the associated duplicate label hidden
'        lblChkGenMin.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenDisableAnim.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenDisableAnim.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenDisableAnim.Visible = False ' make the associated duplicate label hidden
'        lblChkGenDisableAnim.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkIconsZoomOpaque.Enabled = True ' RD does not support storing the configs at the correct location
'        chkIconsZoomOpaque.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblchkIconsZoomOpaque.Visible = False ' make the associated duplicate label hidden
'        lblchkIconsZoomOpaque.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        sliIconsDuration.Enabled = True
'        lblCharacteristicsLabel(6).Enabled = True
'        lblCharacteristicsLabel(11).Enabled = True
'        lblCharacteristicsLabel(12).Enabled = True
'        lblIconsDurationMsCurrent.Enabled = True
'
'        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
''        sliIconsZoomWidth.Enabled = True
''        sliIconsDuration.Enabled = True
'
'        cmbIconsHoverFX.Enabled = True
'
''        Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
''        Call setBounceTypes
'
'        sliAutoHideDuration.Enabled = True
'        sliAnimationInterval.Enabled = True
'
'        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        'fraZoomConfigs.Visible = True
'
'        fraAutoHideDuration.Visible = True
'        fraFontOpacity.Visible = True
'
'
'        optGeneralReadSettings.Enabled = True
'        optGeneralReadRegistry.Enabled = True
'
''        lblGenLabel(0).Enabled = False
''        lblGenLabel(1).Enabled = False
'        sliRunAppInterval.Enabled = False
'        lblGenLabel(2).Enabled = False
'        lblGenRunAppIntervalCur.Enabled = False
'
'        chkGenAlwaysAsk.Enabled = False
'
'        sliAnimationInterval.Enabled = False
'        lblBehaviourLabel(7).Enabled = False
'        lblAnimationIntervalMsLow.Enabled = False
'        lblAnimationIntervalMsHigh.Enabled = False
'        lblAnimationIntervalMsCurrent.Enabled = False
'        lblBehaviourLabel(12).Enabled = False
'
'        cmbAutoHideType.Enabled = False
'
'        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
'
'        imgThemeSample.Enabled = False
'        lblStyleLabel(2).Enabled = False
'
'        sliStyleThemeSize.Enabled = False
'        lblThemeSizeTextHigh.Enabled = False
'        lblStyleSizeCurrent.Enabled = False
'
'        lblBehaviourLabel(5).Enabled = False
'        lblBehaviourLabel(11).Enabled = False
'        sliContinuousHide.Enabled = False
'        lblContinuousHideMsHigh.Enabled = False
'        lblContinuousHideMsCurrent.Enabled = False
'
'        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
'
'        ' RD does not support storing the configs at the correct location
'
'        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        optGeneralReadConfig.Enabled = False ' RD does not support storing the configs at the correct location
'        optGeneralReadConfig.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lbloptGeneralReadConfig.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lbloptGeneralReadConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        optGeneralWriteConfig.Enabled = False ' RD does not support storing the configs at the correct location
'        optGeneralWriteConfig.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblOptGeneralWriteConfig.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblOptGeneralWriteConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkGenAlwaysAsk.Enabled = False ' RD does not support storing the configs at the correct location
'        chkGenAlwaysAsk.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblChkGenAlwaysAsk.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblChkGenAlwaysAsk.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkSplashStatus.Enabled = False ' RD does not support storing the configs at the correct location
'        chkSplashStatus.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblChkSplashStatus.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblChkSplashStatus.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkShowIconSettings.Enabled = False ' RD does not support storing the configs at the correct location
'        chkShowIconSettings.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblGenChkShowIconSettings.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblGenChkShowIconSettings.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkRetainIcons.Enabled = False
'        chkRetainIcons.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblRetainIcons.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblRetainIcons.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'        lblBehaviourLabel(14).Enabled = False
'
'    Else
        ' check where/if steamydock is installed
        Call checkSteamyDockInstallation 'defaultDock is set here
        
        rDDefaultDock = "steamydock"
        defaultDock = 1 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
        
        ' .17 DAEB 07/09/2022 docksettings the dock folder location now changes as it is switched between Rocketdock and Steamy Dock
        dockAppPath = sdAppPath
        txtAppPath.Text = sdAppPath
        
        ' .19 DAEB 07/09/2022 docksettings when you select rocketdock it reverts to the registry but when you select steamydock it does not revert to the dock settings file.
        optGeneralReadConfig.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
        optGeneralWriteConfig.Value = True ' we just want to set this checkbox but we don't want this to trigger a click

        'disable all the controls that steamy dock does not support
                
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkGenMin.Enabled = False ' RD does not support storing the configs at the correct location
        chkGenMin.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        'lblChkGenMin.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        'lblChkGenMin.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        
        'lblChkMinimise.Enabled = False
        'cmbIconActivationFX.Enabled = False
        'cmbStyleTheme.Enabled = False ' does not support themes yet
        'cmbPositionMonitor.Enabled = False
        'cmbIconsQuality.Enabled = False '  does not support enhanced or lower quality icons
        
        sliIconsDuration.Enabled = False ' ' does not support animations at all
        lblCharacteristicsLabel(6).Enabled = False
        lblCharacteristicsLabel(11).Enabled = False
        lblCharacteristicsLabel(12).Enabled = False
        lblIconsDurationMsCurrent.Enabled = False
        
        
        'chkOpenRunning.Enabled = False ' does not support showing opening running applications, always opens new apps.
        'lblChkOpenRunning.Enabled = False
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        sliIconsZoomWidth.Enabled = False ' does not support zoomwidth though this is possible later
'        sliIconsDuration.Enabled = False ' does not support animations at all

        '.nn cmbIconsHoverFX.Enabled = False ' does not support hover effects other than the default
        '.nn sliAutoHideDuration.Enabled = False ' does not support animation at all
        'sliAnimationInterval.Enabled = False ' does not support animation at all
                
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkGenDisableAnim.Enabled = False ' RD does not support storing the configs at the correct location
        chkGenDisableAnim.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        'lblChkGenDisableAnim.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        'lblChkGenDisableAnim.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
                
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkIconsZoomOpaque.Enabled = False ' RD does not support storing the configs at the correct location
        chkIconsZoomOpaque.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblchkIconsZoomOpaque.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblchkIconsZoomOpaque.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' Some of the controls have been bundled onto frames so that they can all be hidden entirely for Steamydock users
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
        'fraZoomConfigs.Visible = False
        
        'fraAutoHideDuration.Visible = true
        fraFontOpacity.Visible = True
        
        optGeneralReadConfig.Enabled = True


        optGeneralWriteConfig.Enabled = True

'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
        
'        If optGeneralReadConfig.Value = True And steamyDockInstalled = True And rocketDockInstalled = True Then
'            chkGenAlwaysAsk.Enabled = True
'            'lblChkAlwaysConfirm.Enabled = True
'        End If

        sliAnimationInterval.Enabled = True
        lblBehaviourLabel(7).Enabled = True
        lblAnimationIntervalMsLow.Enabled = True
        lblAnimationIntervalMsHigh.Enabled = True
        lblAnimationIntervalMsCurrent.Enabled = True
        lblBehaviourLabel(12).Enabled = True
        
        cmbAutoHideType.Enabled = True
        
        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        imgThemeSample.Enabled = True
        lblStyleLabel(2).Enabled = True
        lblStyleLabel(3).Enabled = True
        lblStyleLabel(4).Enabled = True
        lblStyleLabel(5).Enabled = True
        lblStyleLabel(8).Enabled = True
        lblStyleLabel(9).Enabled = True
        lblStyleLabel(10).Enabled = True
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
        
        lblBehaviourLabel(5).Enabled = True
        lblBehaviourLabel(11).Enabled = True
        sliContinuousHide.Enabled = True
        lblContinuousHideMsHigh.Enabled = True
        lblContinuousHideMsCurrent.Enabled = True
        
        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        optGeneralReadConfig.Enabled = True ' RD does not support storing the configs at the correct location
        optGeneralReadConfig.Width = 5820 ' set the width to show the full check box and its intrinsic label
        lbloptGeneralReadConfig.Visible = False ' make the associated duplicate label hidden
        lbloptGeneralReadConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        optGeneralWriteConfig.Enabled = True ' RD does not support storing the configs at the correct location
        optGeneralWriteConfig.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblOptGeneralWriteConfig.Visible = False ' make the associated duplicate label hidden
        'lblOptGeneralWriteConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenAlwaysAsk.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenAlwaysAsk.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenAlwaysAsk.Visible = False ' make the associated duplicate label hidden
'        lblChkGenAlwaysAsk.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkSplashStatus.Enabled = True ' RD does not support storing the configs at the correct location
        chkSplashStatus.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblChkSplashStatus.Visible = False ' make the associated duplicate label hidden
        'lblChkSplashStatus.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
                
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkShowIconSettings.Enabled = True ' RD does not support storing the configs at the correct location
        chkShowIconSettings.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblGenChkShowIconSettings.Visible = False ' make the associated duplicate label hidden
        'lblGenChkShowIconSettings.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkRetainIcons.Enabled = True
        chkRetainIcons.Width = 5820  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        'lblRetainIcons.Visible = False ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        'lblRetainIcons.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        lblBehaviourLabel(14).Enabled = True ' associated title label
        
    'End If
    
    Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
    Call setBounceTypes
    Call populateSoundSelectionDropDown
    Call populateWallpaperStyleDropDown
    Call populateWallpaperTimerIntervalDropDown
    
    chkAutomaticWallpaperChange.Value = CInt(rDAutomaticWallpaperChange)
    chkMoveWinTaskbar.Value = CInt(rDMoveWinTaskbar)
    cmbWallpaperTimerInterval.ListIndex = CInt(rDWallpaperTimerIntervalIndex)

    Call setHidingKey
        
   On Error GoTo 0
   Exit Sub

cmbDefaultDock_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDefaultDock_Click of Form dockSettings"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   If debugflg = 1 Then Debug.Print "%btnFacebook_Click"

    mnuFacebook_Click

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   If debugflg = 1 Then Debug.Print "%btnUpdate_Click"

    mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeMain_MouseDown
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo fmeMain_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%fmeMain_MouseDown"
   


    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    Else
        If Index = 5 Then cmbWallpaper.SetFocus
    End If
    
    ' setting capture of the mouseEnter event on the frame
'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hWnd Then 'MouseEnter
'                Call SetCapture(.hWnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then debug.print "%fmeMain_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If
    
   On Error GoTo 0
   Exit Sub

fmeMain_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeMain_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeSizePreview_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeSizePreview_Click()
   On Error GoTo fmeSizePreview_Click_Error
   If debugflg = 1 Then Debug.Print "%fmeSizePreview_Click"
   
    ' setting capture of the mouseEnter event on the frame
'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hWnd Then 'MouseEnter
'                Call SetCapture(.hWnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then debug.print "%fmeMain_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If

   On Error GoTo 0
   Exit Sub

fmeSizePreview_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeSizePreview_Click of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Form_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%Form_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_MouseMove
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : If the resizing previews are covered by another window then thet are blanked
'             when a mouse enters the form, if the panel is showing the previews are redrawn
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Form_MouseMove_Error
   'If debugflg = 1 Then debug.print "%Form_MouseMove"
   
' setting capture of the mouseEnter event on the form causes weird delays on the whole operation
' of the form controls, so it is now commented out

'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hwnd Then 'MouseEnter
'                Call SetCapture(.hwnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then debug.print "%Form_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If
'    If fmeMain(1).Visible = True Then
'        Call sliIconsSize_Change
'        Call sliIconsZoom_Change
'    End If


    ' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
    fraScrollbarCover.Visible = True

   On Error GoTo 0
   Exit Sub

Form_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Form dockSettings"
End Sub





Private Sub lblAboutPara4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub


Private Sub lblAboutPara5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblPunklabsLink_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblPunklabsLink_Click(Index As Integer)
    Dim answer As VbMsgBoxResult: answer = vbNo
   
    On Error GoTo lblPunklabsLink_Click_Error
   
    If debugflg = 1 Then Debug.Print "%lblPunklabsLink_Click"
    
    ' .22 DAEB 02/10/2022 docksettings added a message pop up on the punklabs link
    answer = MsgBox("This link opens a browser window and connects to Punklabs Homepage. Would you like to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.punklabs.com", vbNullString, App.Path, 1)
    End If
    
   On Error GoTo 0
   Exit Sub

lblPunklabsLink_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblPunklabsLink_Click of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    ' set the menu checks
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Enable"
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Disable"
            themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    
    rDSkinTheme = "dark"
    
    'load the gear images
    imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1.jpg")
    imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3.jpg")

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error

    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    rDSkinTheme = "light"
    
    'load the gear images
    imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1Light.jpg")
    imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3Light.jpg")

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form dockSettings"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(redC As Integer, greenC As Integer, blueC As Integer)
    
        
    ' variables declared
    Dim a As Long
    Dim Ctrl As Control
    Dim useloop As Integer
    
    'initialise the dimensioned variables
     a = 0
     'Ctrl As Control
     useloop = 0
    
    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    ' a method of looping through all the controls that require reversion of any background colouring
    ' note: all buttons need to be style = graphical in order to theme by colour
    For Each Ctrl In dockSettings.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
    Else
        classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
    End If
    
    ' these elements are normal elements that should have their styling reverted
    ' the loop above changes the background colour and we don't want that for all items
    
    iconBox.BackColor = vbWhite
    
    ' loop through the selection icons and revert these to white
    For useloop = 0 To 6
        
        lblText(useloop).BackColor = vbWhite
        'If useloop > 0 Then picIcon(useloop).BackColor = vbWhite
        'lblText(useloop).BackColor = vbWhite
    Next useloop
    
    fmeLblGeneral.BackColor = vbWhite
    'fmeLblFrame(1).BackColor = vbWhite
    fmeLblBehaviour.BackColor = vbWhite

    ' now set the frames that underly the selection icons and revert these to white
    fmeGeneral.BackColor = vbWhite
    fmeIcons.BackColor = vbWhite
    fmeBehaviour.BackColor = vbWhite
    fmeStyle.BackColor = vbWhite
    fmePosition.BackColor = vbWhite
    fmeWallpaper.BackColor = vbWhite
    fmeAbout.BackColor = vbWhite
'    fmeIconBehaviour.BackColor = vbWhite
'    fmeIconAbout.BackColor = vbWhite
'    fmeIconStyle.BackColor = vbWhite
'    fmeIconIcons.BackColor = vbWhite
'    fmeIconPosition.BackColor = vbWhite
'    fmeWallpaper.BackColor = vbWhite
'
    ' labels within the preview box that must stay the high contrast colours
    Label9.BackColor = RGB(212, 208, 199)
    Label13.BackColor = RGB(212, 208, 199)
    Label1.BackColor = RGB(212, 208, 199)
    
        
    lblAboutText.BackColor = RGB(redC, greenC, blueC)
    
    picBusy.BackColor = RGB(redC, greenC, blueC)
    picStylePreview.BackColor = RGB(212, 208, 199)
    picSizePreview.BackColor = RGB(212, 208, 199)
    picZoomSize.BackColor = RGB(212, 208, 199)
    picMinSize.BackColor = RGB(212, 208, 199)
    
    ' now style the reamining elements by hand to the lighter theme colour RGB(redC, greenC, blueC)
    
    'all other buttons go here
    
'    btnGeneralRdFolder.BackColor = RGB(redC, greenC, blueC)
'
'    btnGeneralDockEditor.BackColor = RGB(redC, greenC, blueC)
'    btnGeneralDockSettingsEditor.BackColor = RGB(redC, greenC, blueC)
'    btnGeneralIconSettingsEditor.BackColor = RGB(redC, greenC, blueC)
    
    sliAutoHideDuration.BackColor = RGB(redC, greenC, blueC)
    sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    sliBehaviourAutoHideDelay.BackColor = RGB(redC, greenC, blueC)
    sliBehaviourPopUpDelay.BackColor = RGB(redC, greenC, blueC)
    
    '.0n DAEB Added themeing to two new sliders
    sliStyleThemeSize.BackColor = RGB(redC, greenC, blueC)
    sliStyleFontOpacity.BackColor = RGB(redC, greenC, blueC)
    
    sliContinuousHide.BackColor = RGB(redC, greenC, blueC)
    
    'general tab slider
    sliRunAppInterval.BackColor = RGB(redC, greenC, blueC)

    
    'style tab sliders
    
    sliStyleOpacity.BackColor = RGB(redC, greenC, blueC)
    sliStyleShadowOpacity.BackColor = RGB(redC, greenC, blueC)
    sliStyleOutlineOpacity.BackColor = RGB(redC, greenC, blueC)
    
    'position tab sliders
    
    sliPositionCentre.BackColor = RGB(redC, greenC, blueC)
    sliPositionEdgeOffset.BackColor = RGB(redC, greenC, blueC)
    
    ' icons tab picture and frame elements
        
'    picSizePreview.BackColor = RGB(redC, greenC, blueC)
'    Label9.BackColor = RGB(redC, greenC, blueC)
'    Label13.BackColor = RGB(redC, greenC, blueC)
'    Label1
    
    ' icons tab picboxes
    
'    picZoomSize.BackColor = RGB(redC, greenC, blueC)
'    picMinSize.BackColor = RGB(redC, greenC, blueC)
'    picHiddenPicture.BackColor = RGB(redC, greenC, blueC)
    
    ' icons tab sliders
    
    sliIconsOpacity.BackColor = RGB(redC, greenC, blueC)
    sliIconsSize.BackColor = RGB(redC, greenC, blueC)
    sliIconsZoom.BackColor = RGB(redC, greenC, blueC)
    sliIconsZoomWidth.BackColor = RGB(redC, greenC, blueC)
    sliIconsDuration.BackColor = RGB(redC, greenC, blueC)

    
    PutINISetting "Software\DockSettings", "SkinTheme", rDSkinTheme, toolSettingsFile ' now saved to the toolsettingsfile

    End Sub
'---------------------------------------------------------------------------------------
' Procedure : picCogs1_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picCogs1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo picCogs1_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%picCogs1_MouseDown"
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

picCogs1_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picCogs1_MouseDown of Form dockSettings"
End Sub

Private Sub optGeneralReadSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadSettings.hWnd, "This option allows you to read the configuration from Rocketdock's program files folder, this is for migrating in a read-only fashion from RocketDock to SteamyDock. Requires admin access so only select this option when migrating from Rocketdock. ", _
                  TTIconInfo, "Help on reading from the settings.ini.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteConfig_Click
' Author    : beededea
' Date      : 05/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralWriteConfig_Click()

   On Error GoTo optGeneralWriteConfig_Click_Error

    If startupFlg = True Then '
        ' don't do this on the first startup run
        Exit Sub
    Else
    
        rDGeneralWriteConfig = optGeneralWriteConfig.Value ' this is the nub


'        If optGeneralWriteConfig.Value = True Then
'            rDGeneralWriteConfig = "True"
'
'        Else
'            rDGeneralWriteConfig = "False"
'        End If
    
    End If

   On Error GoTo 0
   Exit Sub

optGeneralWriteConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteConfig_Click of Form dockSettings"

End Sub

Private Sub optGeneralWriteConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteConfig.hWnd, "This option stores ALL configuration within the user data area retaining future compatibility in Windows. Not available to Rocketdock.", _
                  TTIconInfo, "Help on Writing SteamyDock's Config.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteRegistry_Click
' Author    : beededea
' Date      : 05/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub optGeneralWriteRegistry_Click()
'   On Error GoTo optGeneralWriteRegistry_Click_Error
'
'    If optGeneralWriteRegistry.Value = True Then
'        ' nothing to do, the checkbox value is used later to determine where to write the data
'    End If
'    If defaultDock = 0 Then optGeneralReadRegistry.Value = True ' if running Rocketdock the two must be kept in sync
'
'    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location
'
'   On Error GoTo 0
'   Exit Sub
'
'optGeneralWriteRegistry_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteRegistry_Click of Form dockSettings"
'End Sub

'Private Sub optGeneralWriteRegistry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteRegistry.hWnd, "Stores the configuration in the Rocketdock portion of the Registry, incompatible with newer version of Windows, this can cause some security problems and in all case requires admin rights to operate. Best to use option 3. ", _
'                  TTIconInfo, "Help on writing settings to the registry.", , , , True
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteSettings_Click
' Author    : beededea
' Date      : 01/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub optGeneralWriteSettings_Click()
'
'   On Error GoTo optGeneralWriteSettings_Click_Error
'
'    tmpSettingsFile = rdAppPath & "\tmpSettings.ini" ' temporary copy of Rocketdock 's settings file
'
'    If startupFlg = True Then '
'        ' don't do this on the first startup run
'        Exit Sub
'    Else
'
'        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
'            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
'            ' create a settings.ini file in the rocketdock folder
'            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
'            Close #1         ' close the file and
'             ' test it exists
'            If fFExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
'                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
'                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
'            End If
'        End If
'
'        If defaultDock = 0 Then ' Rocketdock
'            If optGeneralWriteSettings.Value = True Then ' keep the two in synch.
'                If optGeneralReadSettings.Value = False Then
'                    optGeneralReadSettings.Value = True
'                End If
'            End If
'        End If
'    End If
'
'    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location
'
'   On Error GoTo 0
'   Exit Sub
'
'optGeneralWriteSettings_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteSettings_Click of Form dockSettings"
'
'End Sub

'Private Sub optGeneralWriteSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteSettings.hWnd, "Store configuration in Rocketdock's program files folder, can cause security issues on newer systems beyond XP and requires admin access. Best to move to option 3. ", _
'                  TTIconInfo, "Help on Storing within Rocketdock's program files folder.", , , , True
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : imgMultipleGears3_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub imgMultipleGears3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo imgMultipleGears3_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%imgMultipleGears3_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

imgMultipleGears3_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgMultipleGears3_MouseDown of Form dockSettings"
End Sub







    
'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
        
    ' variables declared
    Dim toolSettingsDir As String
    
    'initialise the dimensioned variables
    toolSettingsDir = vbNullString
    
    On Error GoTo getToolSettingsFile_Error
    If debugflg = 1 Then Debug.Print "%getToolSettingsFile"
    
    toolSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\dockSettings" ' just for this user alone
    toolSettingsFile = toolSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(toolSettingsDir) Then
        MkDir toolSettingsDir
    End If
    
    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(toolSettingsFile) Then
        FileCopy App.Path & "\settings.ini", toolSettingsFile
    End If
    
    'confirm the settings file exists, if not use the version in the app itself
    If Not fFExists(toolSettingsFile) Then
        toolSettingsFile = App.Path & "\settings.ini"
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form dockSettings"

End Sub
    



'---------------------------------------------------------------------------------------
' Procedure : readRocketdockSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRocketdockSettings()

'    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    
    ' the first is the RD settings file that only exists if RD is NOT using the registry
    ' the second is the settings file for this tool to store its own preferences
        
    ' check to see if the first settings file exists
    
    On Error GoTo readRocketdockSettings_Error
   

    If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
        If optGeneralReadConfig.Value = False Then
'                optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'                optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
        End If
        ' here we read from the settings file
        readDockSettingsFile "Software\RocketDock", origSettingsFile
        Call validateInputs
    Else
        If optGeneralReadConfig.Value = False Then
            optGeneralReadRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
            'optGeneralWriteRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
        End If

        ' read the dock configuration from the registry into variables
        Call readRegistry
    End If


    
    On Error GoTo 0
    Exit Sub

readRocketdockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRocketdockSettings of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readAndSetUtilityFont
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : reads the tool's font settings from the local tool file
'---------------------------------------------------------------------------------------
'
Private Sub readAndSetUtilityFont()
  
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedStyle As String
    'Dim suppliedColour As Variant
    
    'initialise the dimensioned variables
    suppliedFont = vbNullString
    suppliedSize = 0
    suppliedWeight = 0
    suppliedStyle = False
    'suppliedColour = Empty

    On Error GoTo readAndSetUtilityFont_Error
    
    ' set the tool's default font
    suppliedFont = GetINISetting("Software\DockSettings", "defaultFont", toolSettingsFile)
    suppliedSize = Val(GetINISetting("Software\DockSettings", "defaultSize", toolSettingsFile))
    suppliedWeight = Val(GetINISetting("Software\DockSettings", "defaultStrength", toolSettingsFile))
    suppliedStyle = GetINISetting("Software\DockSettings", "defaultStyle", toolSettingsFile)
    rDSkinTheme = GetINISetting("Software\DockSettings", "SkinTheme", toolSettingsFile)
    
    If suppliedSize = 0 Then suppliedSize = 8
    gblSuppliedFont = suppliedFont
    gblSuppliedFontSize = suppliedSize
        
    If Not suppliedFont = vbNullString Then
        Call changeFont(suppliedFont, suppliedSize, suppliedWeight, CBool(LCase(suppliedStyle)))
    End If

   On Error GoTo 0
   Exit Sub

readAndSetUtilityFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readAndSetUtilityFont of Form dockSettings on line " & Erl
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPreviewFontColours
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPreviewFontColours(suppliedColour)
   On Error GoTo setPreviewFontColours_Error
   If debugflg = 1 Then Debug.Print "%setPreviewFontColours"

    lblPreviewFont.ForeColor = suppliedColour

   On Error GoTo 0
   Exit Sub

setPreviewFontColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPreviewFontColours of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : setPreviewConvertedFontColours
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPreviewConvertedFontColours(suppliedColour)
        
   On Error GoTo setPreviewConvertedFontColours_Error
   If debugflg = 1 Then Debug.Print "%setPreviewConvertedFontColours"

    lblPreviewFont.ForeColor = Convert_Dec2RGB(suppliedColour)

   On Error GoTo 0
   Exit Sub

setPreviewConvertedFontColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPreviewConvertedFontColours of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : writeRegistry
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : utility needs admin to write to Rocketdock's registry entries
'---------------------------------------------------------------------------------------
'
'Private Sub writeRegistry()
'
'    On Error GoTo writeRegistry_Error
'
'    ' all tested and working but ONLY when run as admin
'
'    'general panel
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", rDLockIcons)
'    ' rDRetainIcons not required
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OpenRunning", rDOpenRunning)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ShowRunning", rDShowRunning)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ManageWindows", rDManageWindows)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "DisableMinAnimation", rDDisableMinAnimation)
'
'    'icon panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconQuality", Val(rDIconQuality))
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconOpacity", rDIconOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomOpaque", rDZoomOpaque)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMin", rDIconMin)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HoverFX", rDHoverFX)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMax", rdIconMax)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomWidth", Val(rDZoomWidth))
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomTicks", rDZoomTicks)
'
'    'behaviour panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconActivationFX", rDIconActivationFX)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHide", rDAutoHide) '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideTicks", rDAutoHideDuration)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideDelay", rDAutoHideDelay)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "MouseActivate", rDMouseActivate)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "PopupDelay", rDPopupDelay)
'
'
'    'position panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Monitor", rDMonitor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", rDSide)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "zOrderMode", rDzOrderMode)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Offset", rDOffset)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "vOffset", rDvOffset)
'
'    'style panel
'    'If rDtheme = "blank" Then rDtheme = ""
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "theme", rDtheme)
'    'If rDtheme = "" Then rDtheme = "blank"
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ThemeOpacity", rDThemeOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HideLabels", rDHideLabels)
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontName", rDFontName) '*
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontColor", rDFontColor) '*
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontSize", rDFontSize)
'    'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontCharSet", rD)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontFlags", rDFontFlags) '*
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowColor", rDFontShadowColor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineColor", rDFontOutlineColor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineOpacity", rDFontOutlineOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowOpacity", rDFontShadowOpacity)
'
'   On Error GoTo 0
'   Exit Sub
'
'writeRegistry_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistry of Form dockSettings"
'End Sub




'Private Sub picIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip picIcon.hWnd, "This button opens the panel to configure the general options that apply to the whole dock program.", _
'                  TTIconInfo, "Help on the General Options Button", , , , True
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : imgIcon_MouseDown
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo imgIcon_MouseDown_Error

    Call imgIcon_MouseDown_Event(Index)

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

imgIcon_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgIcon_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : imgIcon_MouseDown_Event
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub imgIcon_MouseDown_Event(Index As Integer)
   On Error GoTo imgIcon_MouseDown_Event_Error

    rDOptionsTabIndex = CStr(Index + 1)
    If Index < 0 Then Index = 0
    
    'CFG - write the current open tab to the 3rd config settings
    ' .20 DAEB 07/09/2022 docksettings tab selection fixed
    PutINISetting "Software\DockSettings", "OptionsTabIndex", rDOptionsTabIndex, toolSettingsFile
    
    If Index <> 0 Then fmeMain(0).Visible = False
    If Index <> 1 Then fmeMain(1).Visible = False
    If Index <> 2 Then fmeMain(2).Visible = False
    If Index <> 3 Then fmeMain(3).Visible = False
    If Index <> 4 Then fmeMain(4).Visible = False
    If Index <> 5 Then fmeMain(5).Visible = False
    If Index <> 6 Then fmeMain(6).Visible = False
    
    fmeMain(Index).Visible = True

    fmeMain(Index).Left = 1665 * gblResizeRatio
    fmeMain(Index).top = 30 * gblResizeRatio
    
    ' ensure the resizing icons always display.
    ' it seems that when the picturebox is hidden and given focus then the images are lost.
    ' calling these routines restores the images.
        
    imgIcon(Index).Visible = False
    imgIconPressed(Index).Visible = True
    
    If Index = 5 And dockSettings.Visible = True Then
        cmbWallpaper.SetFocus
    End If
    
    If Index = 6 And dockSettings.Visible = True Then
        lblAboutText.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

imgIcon_MouseDown_Event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgIcon_MouseDown_Event of Form dockSettings"

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : fmeicon_MouseMove
'' Author    : beededea
'' Date      : 31/03/2025
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub fmeicon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim descriptiveText As String
'    Dim titleText As String
'
'   On Error GoTo fmeicon_MouseMove_Error
'
'    descriptiveText = ""
'    titleText = ""
'
'    If rDEnableBalloonTooltips = "1" Then
'        If Index = 0 Then
'            descriptiveText = "This Button will select the general pane. Use this panel to configure the general options that apply to the whole dock program. "
'            titleText = "Help on the General Pane Button."
'        ElseIf Index = 1 Then
'            descriptiveText = "This Button will select the characteristics pane. Use this panel to configure the icon characteristics that apply only to the icons themselves. "
'            titleText = "Help on the Icon Characteristics Pane Button."
'        ElseIf Index = 2 Then
'            descriptiveText = "This Button will select the behaviour pane. Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
'            titleText = "Help on the Behaviour Pane Button."
'        ElseIf Index = 3 Then
'            descriptiveText = "This Button will select the style, themes and fonts pane. This is used to configure the label and font settings."
'            titleText = "Help on the Style Themes and Fonts Pane Button."
'        ElseIf Index = 4 Then
'            descriptiveText = "This Button will select the position pane. This pane is used to control the location of the dock. "
'            titleText = "Help on the Position Pane Button."
'        ElseIf Index = 5 Then
'            descriptiveText = "This Button will select the wallpaper pane. The wallpaper Panel allows you to select and apply a background image as the desktop wallpaper."
'            titleText = "Help on the Wallpaper Pane Button."
'        ElseIf Index = 6 Then
'            descriptiveText = "This Button will select the about pane. The Position Panel provides the version number of this utility, useful information when reporting a bug. The text below this gives due credit to Punk labs for being the originator of  and gives thanks to them for coming up with such a useful tool and also to Apple who created the original idea for this whole genre of docks. This pane also gives access to some useful utilities."
'            titleText = "Help on the Position Pane Button."
'        End If
'    End If
'
'    CreateToolTip fmeicon(Index).hWnd, descriptiveText, TTIconInfo, titleText, , , , True
'
'   On Error GoTo 0
'   Exit Sub
'
'fmeicon_MouseMove_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeicon_MouseMove of Form dockSettings"
'
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : fmeLblBehaviour_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeLblBehaviour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeLblBehaviour_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the behaviour pane. Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
        titleText = "Help on the Behaviour Pane Button."
        CreateToolTip fmeLblBehaviour.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeLblBehaviour_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeLblBehaviour_MouseMove of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeIcons_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeIcons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeIcons_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
            descriptiveText = "This Button will select the icon characteristics pane. Use this panel to configure the icon characteristics that apply only to the icons themselves. "
            titleText = "Help on the Icon Characteristics Pane Button."
        CreateToolTip fmeIcons.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeIcons_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeIcons_MouseMove of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fmeGeneral_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeGeneral_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the general pane. Use this panel to configure the general options that apply to the whole dock program. "
        titleText = "Help on the General Pane Button."
        CreateToolTip fmeGeneral.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeGeneral_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeGeneral_MouseMove of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fmeLblGeneral_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeLblGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeLblGeneral_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the general pane. Use this panel to configure the general options that apply to the whole dock program. "
        titleText = "Help on the General Pane Button."
        CreateToolTip fmeLblGeneral.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeLblGeneral_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeLblGeneral_MouseMove of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmePosition_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmePosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmePosition_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the position pane. This pane is used to control the location of the dock. "
        titleText = "Help on the Position Pane Button."
        CreateToolTip fmePosition.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmePosition_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmePosition_MouseMove of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fmeAbout_MouseMove
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim descriptiveText As String: descriptiveText = vbNullString
    Dim titleText As String: titleText = vbNullString

    On Error GoTo fmeAbout_MouseMove_Error

    If rDEnableBalloonTooltips = "1" Then
        descriptiveText = "This Button will select the about pane. The About Panel provides the version number of this utility, useful information when reporting a bug. The text below this gives due credit to Punk labs for being the originator of  and gives thanks to them for coming up with such a useful tool and also to Apple who created the original idea for this whole genre of docks. This pane also gives access to some useful utilities."
        titleText = "Help on the About Pane Button."
        CreateToolTip fmeAbout.hWnd, descriptiveText, TTIconInfo, titleText, , , , True
    End If

   On Error GoTo 0
   Exit Sub

fmeAbout_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeAbout_MouseMove of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : imgIcon_MouseUp
' Author    : beededea
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
   On Error GoTo imgIcon_MouseUp_Error

    imgIcon(Index).Visible = True
    imgIconPressed(Index).Visible = False

   On Error GoTo 0
   Exit Sub

imgIcon_MouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgIcon_MouseUp of Form dockSettings"
End Sub



Private Sub picMinSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picMinSize.hWnd, "This frame shows the icon in the small size just as it will look in the dock.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True
End Sub

Private Sub picSizePreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picSizePreview.hWnd, "This frame shows the icon in two sizes, as it looks in the dock (on the left) and is it will appear when fully enlarged during a zoom.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picStylePreview_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picStylePreview_Click()
       
    ' variables declared
    Dim colourResult As Long
        
    'initialise the dimensioned variables
    colourResult = 0
    
    On Error GoTo picStylePreview_Click_Error

    colourResult = ShowColorDialog(Me.hWnd, True, rDFontShadowColor)

    If colourResult <> -1 And colourResult <> 0 Then
        picStylePreview.BackColor = colourResult
    End If

   On Error GoTo 0
   Exit Sub

picStylePreview_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picStylePreview_Click of Form dockSettings"
End Sub

Private Sub picStylePreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picStylePreview.hWnd, "This panel shows a preview of the font selection - you can change the background of the preview to approximate how your font will look  on your desktop.", _
                  TTIconInfo, "Help on the Font Preview Pane.", , , , True
End Sub
'Private Sub imgThemeSample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip imgThemeSample.hWnd, "This panel shows a portion of the dock with the current theme selected.", _
'                  TTIconInfo, "Help on Theme Selection.", , , , True
'End Sub

Private Sub picZoomSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picZoomSize.hWnd, "This frame shows the icon in the large size just as it looks when fully enlarged during a mouse-over zoom.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : repaintTimer_Timer
' Author    : beededea
' Date      : 09/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub repaintTimer_Timer()

   On Error GoTo repaintTimer_Timer_Error
   If debugflg = 1 Then Debug.Print "%repaintTimer_Timer"
    If fmeMain(1).Visible = True Then
        Call sliIconsSize_Change
        Call sliIconsZoom_Change
    End If
   On Error GoTo 0
   Exit Sub

repaintTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure repaintTimer_Timer of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliAnimationInterval_Change
' Author    : beededea
' Date      : 10/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliAnimationInterval_Change()

   On Error GoTo sliAnimationInterval_Change_Error
    lblAnimationIntervalMsCurrent.Caption = "(" & sliAnimationInterval.Value & ")"

    rDAnimationInterval = sliAnimationInterval.Value


   On Error GoTo 0
   Exit Sub

sliAnimationInterval_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliAnimationInterval_Change of Form dockSettings"
End Sub

Private Sub sliAnimationInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliAnimationInterval.hWnd, "The overall animation period in millisecs. 10ms is a good default but experiment with the value for your own system if the animation is not as smooth as you desire. The animation is achieved using GDI+ and is entirely CPU driven. You may see a benefit in Steamydock by changing this slider. This will have no effect on Rocketdock.", _
                  TTIconInfo, "Help on the Animation Interval.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliBehaviourAutoHideDelay_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliBehaviourAutoHideDelay_Change()
   On Error GoTo sliBehaviourAutoHideDelay_Click_Error
   If debugflg = 1 Then Debug.Print "%sliBehaviourAutoHideDelay_Click"

    lblAutoHideDelayMsCurrent.Caption = "(" & 3 + (sliBehaviourAutoHideDelay.Value / 1000) & ") secs"
    
    rDAutoHideDelay = sliBehaviourAutoHideDelay.Value

   On Error GoTo 0
   Exit Sub

sliBehaviourAutoHideDelay_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliBehaviourAutoHideDelay_Click of Form Form1"
End Sub

Private Sub sliBehaviourAutoHideDelay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliBehaviourAutoHideDelay.hWnd, "Determine the delay between the last usage of the dock and when it will auto-hide.", _
                  TTIconInfo, "Help on the AutoHide Delay Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliAutoHideDuration_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliAutoHideDuration_Change()
   On Error GoTo sliAutoHideDuration_Change_Error
   If debugflg = 1 Then Debug.Print "%sliAutoHideDuration_Change"

    lblAutoHideDurationMsCurrent.Caption = "(" & sliAutoHideDuration.Value & ")"

    rDAutoHideDuration = sliAutoHideDuration.Value

   On Error GoTo 0
   Exit Sub

sliAutoHideDuration_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliAutoHideDuration_Change of Form Form1"
End Sub

Private Sub sliAutoHideDuration_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliAutoHideDuration.hWnd, "The speed at which the dock auto-hide animation will occur.", _
                  TTIconInfo, "Help on the AutoHide Duration Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliBehaviourPopUpDelay_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliBehaviourPopUpDelay_Change()
   On Error GoTo sliBehaviourPopUpDelay_Change_Error
   If debugflg = 1 Then Debug.Print "%sliBehaviourPopUpDelay_Change"

    lblBehaviourPopUpDelayMsCurrrent.Caption = "(" & sliBehaviourPopUpDelay.Value & ")"
    
    rDPopupDelay = sliBehaviourPopUpDelay.Value

   On Error GoTo 0
   Exit Sub

sliBehaviourPopUpDelay_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliBehaviourPopUpDelay_Change of Form Form1"
End Sub



Private Sub sliBehaviourPopUpDelay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliBehaviourPopUpDelay.hWnd, "The speed at which the dock auto-reveal animation will occur. This was previously called the Pop-up Delay in Rocketdock's settings screen.", _
                  TTIconInfo, "Help on the AutoReveal Duration Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliContinuousHide_Change
' Author    : beededea
' Date      : 25/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliContinuousHide_Change()

   On Error GoTo sliContinuousHide_Change_Error

    If sliContinuousHide.Value = 1 Then
        lblContinuousHideMsCurrent.Caption = "(" & sliContinuousHide.Value & ") min"
    Else
        lblContinuousHideMsCurrent.Caption = "(" & sliContinuousHide.Value & ") mins"
    End If
    sDContinuousHide = sliContinuousHide.Value

   On Error GoTo 0
   Exit Sub

sliContinuousHide_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliContinuousHide_Change of Form dockSettings"
    
End Sub

Private Sub sliContinuousHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliContinuousHide.hWnd, "Determine the amount of time the dock will disappear when told to go away using F11 key.", _
                  TTIconInfo, "Help on the Continuous Hide Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliRunAppInterval_Change
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliRunAppInterval_Change()

   On Error GoTo sliRunAppInterval_Change_Error

    lblGenRunAppIntervalCur.Caption = "(" & sliRunAppInterval.Value & " seconds)"
    rDRunAppInterval = sliRunAppInterval.Value

   On Error GoTo 0
   Exit Sub

sliRunAppInterval_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliRunAppInterval_Change of Form dockSettings"
    
End Sub

Private Sub sliRunAppInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliRunAppInterval.hWnd, "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency set here. The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds. ", _
                  TTIconInfo, "Help on the Running Application Timer.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsDuration_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsDuration_Change()
   On Error GoTo sliIconsDuration_Change_Error
   If debugflg = 1 Then Debug.Print "%sliIconsDuration_Change"

    lblIconsDurationMsCurrent.Caption = "(" & sliIconsDuration.Value & "ms)"
    
    rDZoomTicks = sliIconsDuration.Value

   On Error GoTo 0
   Exit Sub

sliIconsDuration_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsDuration_Change of Form Form1"
End Sub

Private Sub sliIconsDuration_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsDuration.hWnd, "How long the effect is applied in milliseconds. ", _
                  TTIconInfo, "Help on the Icon Zoom Duration Slider", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsOpacity_Change()
   On Error GoTo sliIconsOpacity_Change_Error
   If debugflg = 1 Then Debug.Print "%sliIconsOpacity_Change"

    lblIconsOpacity.Caption = "(" & sliIconsOpacity.Value & "%)"
    
    rDIconOpacity = sliIconsOpacity.Value

   On Error GoTo 0
   Exit Sub

sliIconsOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsOpacity_Change of Form Form1"
End Sub

Private Sub sliIconsOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsOpacity.hWnd, "The icons in the dock can be made transparent here.", _
                  TTIconInfo, "Help on the Icon Opacity Slider", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsSize_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsSize_Change()
        
    ' variables declared
    Dim newSize As Integer
        
    'initialise the dimensioned variables
    newSize = 0

    On Error GoTo sliIconsSize_Change_Error
    If debugflg = 1 Then Debug.Print "%sliIconsSize_Change"

    lblIconsSize.Caption = "(" & sliIconsSize.Value & "px)"
          
    newSize = PixelsToTwips(sliIconsSize.Value)
    picMinSize.Cls
    Call picMinSize.PaintPicture(picHiddenPicture, 60 + (1920 / 2) - (newSize / 2), 60 + (1920 / 2) - (newSize / 2), newSize, newSize)

    rDIconMin = sliIconsSize.Value

   On Error GoTo 0
   Exit Sub

sliIconsSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsSize_Change of Form Form1"
End Sub

Private Sub sliIconsSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsSize.hWnd, "The size of all the icons in the dock prior to any zoom effect being applied. ", _
                  TTIconInfo, "Help on the Icons Size Slider", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsZoom_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsZoom_Change()
       
    ' variables declared
    Dim newSize As Long
        
    'initialise the dimensioned variables
    newSize = 0
    
   On Error GoTo sliIconsZoom_Change_Error
   If debugflg = 1 Then Debug.Print "%sliIconsZoom_Change"

    lblIconsZoom.Caption = "(" & sliIconsZoom.Value & "px)"
    
    
    Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation
    
    newSize = PixelsToTwips(sliIconsZoom.Value)
    picZoomSize.Cls
    Call picZoomSize.PaintPicture(picHiddenPicture, 60 + (3840 / 2) - (newSize / 2), 60 + (3840 / 2) - (newSize / 2), newSize, newSize)
    
    
'    Call picZoomSize.PaintPicture(picHiddenPicture, 60 + (1920 / 2) - (newSize / 2), 60 + (1920 / 2) - (newSize / 2), newSize, newSize)
'    Call picZoomSize.PaintPicture(picHiddenPicture, 60, 60, newSize, newSize)
'
'    'picZoomSize.Left = 2640 + (3840 / 2) - (3840 / 2)
'    picZoomSize.Top = 100 + (3840 - newSize)

    rdIconMax = sliIconsZoom.Value



    
    

   On Error GoTo 0
   Exit Sub

sliIconsZoom_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsZoom_Change of Form Form1"
End Sub

   '---------------------------------------------------------------------------------------
    ' Procedure : vb6TwipsToPixels
    ' Author    : beededea
    ' Date      : 17/10/2019
    ' Purpose   : VB6 polyfills, not using VB6 compatibility mode
    ' doing away with VB6 compatibility mode will remove the 32bit limitation...
    '---------------------------------------------------------------------------------------
    '
    Public Function TwipsToPixels(ByVal intTwips As Integer) As Integer
                
    ' variables declared
    Dim nTwips As Integer
   
    'initialise the dimensioned variables
    nTwips = 0

            'vb6TwipsToPixels = intTwips * g.DpiX / 1440
            nTwips = intTwips / Screen.TwipsPerPixelX

            TwipsToPixels = nTwips
    End Function

'---------------------------------------------------------------------------------------
' Procedure : PixelsToTwips
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : RD works with pixels but VB6 works with twips
'---------------------------------------------------------------------------------------
'
Public Function PixelsToTwips(ByVal intPixels As Integer) As Integer

        
    ' variables declared
    Dim nTwips As Integer
        
    'initialise the dimensioned variables
    nTwips = 0
    
   On Error GoTo PixelsToTwips_Error
   If debugflg = 1 Then Debug.Print "%PixelsToTwips"

            'vb6PixelsToTwips = intPixels / g.DpiX * 1440
            nTwips = intPixels * Screen.TwipsPerPixelX
            
            PixelsToTwips = nTwips

   On Error GoTo 0
   Exit Function

PixelsToTwips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PixelsToTwips of Form dockSettings"

End Function

Private Sub sliIconsZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsZoom.hWnd, "The maximum icon size after a zoom. ", _
                  TTIconInfo, "Help on the Icon Zoom Slider", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsZoomWidth_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsZoomWidth_Change()
   On Error GoTo sliIconsZoomWidth_Change_Error
   If debugflg = 1 Then Debug.Print "%sliIconsZoomWidth_Change"

    lblIconsZoomWidth.Caption = "(" & sliIconsZoomWidth.Value & ")"
    
    rDZoomWidth = sliIconsZoomWidth.Value

   On Error GoTo 0
   Exit Sub

sliIconsZoomWidth_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsZoomWidth_Change of Form Form1"
End Sub



Private Sub sliIconsZoomWidth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsZoomWidth.hWnd, "How many icons to the left and right are also animated. Lower power machines will benefit from a lower setting. 4 is fine. ", _
                  TTIconInfo, "Help on the Icon Zoom Width Slider", , , , True


End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliPositionCentre_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliPositionCentre_Change()
   On Error GoTo sliPositionCentre_Change_Error
   If debugflg = 1 Then Debug.Print "%sliPositionCentre_Change"

    lblPositionCentrePercCurrent.Caption = "(" & Val(sliPositionCentre.Value) & "%)"
    
    rDOffset = sliPositionCentre.Value

   On Error GoTo 0
   Exit Sub

sliPositionCentre_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPositionCentre_Change of Form Form1"
End Sub

Private Sub sliPositionCentre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliPositionCentre.hWnd, "You can align the dock so that it is centred or offset as you require.", _
                  TTIconInfo, "Help on the Dock Centre Position Slider ", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliPositionEdgeOffset_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliPositionEdgeOffset_Click()
   On Error GoTo sliPositionEdgeOffset_Click_Error
   If debugflg = 1 Then Debug.Print "%sliPositionEdgeOffset_Click"

    lblPositionEdgeOffsetPxCurrent.Caption = "(" & Val(sliPositionEdgeOffset.Value) & "px)"
    
    rDvOffset = sliPositionEdgeOffset.Value
    
   On Error GoTo 0
   Exit Sub

sliPositionEdgeOffset_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPositionEdgeOffset_Click of Form Form1"
End Sub

Private Sub sliPositionEdgeOffset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliPositionEdgeOffset.hWnd, "Position from the bottom/top edge of the screen.", _
                  TTIconInfo, "Help on the Dock Position Edge Offset Slider ", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleFontOpacity_Click
' Author    : beededea
' Date      : 17/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleFontOpacity_Click()

   On Error GoTo sliStyleFontOpacity_Click_Error

    lblStyleFontOpacityCurrent.Caption = "(" & Val(sliStyleFontOpacity.Value) & "%)"
    
    sDFontOpacity = sliStyleFontOpacity.Value
    
   On Error GoTo 0
   Exit Sub

sliStyleFontOpacity_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleFontOpacity_Click of Form dockSettings"
End Sub

Private Sub sliStyleFontOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleFontOpacity.hWnd, "The font transparency can be changed here.", _
                  TTIconInfo, "Help on the Font Opacity Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleOpacity_Change()
   On Error GoTo sliStyleOpacity_Change_Error
   If debugflg = 1 Then Debug.Print "%sliStyleOpacity_Change"

    lblStyleOpacityCurrent.Caption = "(" & Val(sliStyleOpacity.Value) & "%)"
    
    rDThemeOpacity = sliStyleOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleOpacity_Change of Form Form1"
End Sub

Private Sub sliStyleOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleOpacity.hWnd, "This controls the transparency of the background theme.", _
                  TTIconInfo, "Help on the Opacity Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleOutlineOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleOutlineOpacity_Change()
   On Error GoTo sliStyleOutlineOpacity_Change_Error
   If debugflg = 1 Then Debug.Print "%sliStyleOutlineOpacity_Change"

    lblStyleOutlineOpacityCurrent.Caption = "(" & Val(sliStyleOutlineOpacity.Value) & "%)"

    rDFontOutlineOpacity = sliStyleOutlineOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleOutlineOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleOutlineOpacity_Change of Form Form1"
End Sub

Private Sub sliStyleOutlineOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleOutlineOpacity.hWnd, "The label outline transparency, use the slider to change.", _
                  TTIconInfo, "Help on the Outline Opacity Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleShadowOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleShadowOpacity_Change()
   On Error GoTo sliStyleShadowOpacity_Change_Error
   If debugflg = 1 Then Debug.Print "%sliStyleShadowOpacity_Change"

    lblStyleShadowOpacityCurrent.Caption = "(" & Val(sliStyleShadowOpacity.Value) & "%)"
    
    rDFontShadowOpacity = sliStyleShadowOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleShadowOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleShadowOpacity_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFont_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFont_Click()

        
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As String
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedStyle As Boolean
    Dim suppliedColour As Variant
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    Dim fontSelected As Boolean
        
    'initialise the dimensioned variables
    
    suppliedFont = vbNullString
    suppliedSize = 0
    suppliedWeight = 0
    suppliedStyle = False
    suppliedColour = Empty
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False
    fontSelected = False
    
    On Error GoTo mnuFont_Click_Error
    If debugflg = 1 Then Debug.Print "%mnuFont_Click"

    displayFontSelector suppliedFont, Val(suppliedSize), suppliedWeight, suppliedStyle, suppliedColour, suppliedItalics, suppliedUnderline, fontSelected
    If fontSelected = False Then Exit Sub

    If fFExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\DockSettings", "defaultFont", suppliedFont, toolSettingsFile
        PutINISetting "Software\DockSettings", "defaultSize", suppliedSize, toolSettingsFile
        PutINISetting "Software\DockSettings", "defaultStrength", suppliedWeight, toolSettingsFile
        PutINISetting "Software\DockSettings", "defaultStyle", suppliedStyle, toolSettingsFile
    End If

    If suppliedWeight > 700 Then
        suppliedBold = True
    Else
        suppliedBold = False
    End If
    
    If suppliedFont <> vbNullString Then
        Call changeFont(suppliedFont, Val(suppliedSize), suppliedWeight, suppliedStyle)
    End If

   On Error GoTo 0
   Exit Sub

mnuFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFont_Click of Form dockSettings"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub changeFont(suppliedFont As String, suppliedSize As Integer, suppliedWeight As Integer, suppliedStyle As Boolean)
        
    ' variables declared
    Dim useloop As Integer
    Dim Ctrl As Control
        
    'initialise the dimensioned variables
    useloop = 0
    
    On Error GoTo changeFont_Error
    
    If debugflg = 1 Then Debug.Print "%" & "changeFont"
      
    ' a method of looping through all the controls and identifying the labels and text boxes
    For Each Ctrl In dockSettings.Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
            If Ctrl.Name <> "lblDragCorner" Then
                If suppliedFont <> vbNullString Then Ctrl.Font.Name = suppliedFont
                If suppliedSize > 0 Then Ctrl.Font.Size = suppliedSize
                'Ctrl.Font.Italic = fntItalics
            End If
        End If
    Next
    
    ' The comboboxes all autoselect when the font is changed, we need to reset this afterwards

    cmbIconsQuality.SelLength = 0
    cmbIconsHoverFX.SelLength = 0
    'cmbDefaultDock.SelLength = 0
    cmbHidingKey.SelLength = 0
    cmbDefaultDock.SelLength = 0
    cmbIconActivationFX.SelLength = 0
    cmbBehaviourSoundSelection.SelLength = 0
    cmbStyleTheme.SelLength = 0
    cmbPositionMonitor.SelLength = 0
    cmbPositionScreen.SelLength = 0
    cmbPositionLayering.SelLength = 0
    cmbAutoHideType.SelLength = 0
    cmbWallpaper.SelLength = 0
    cmbWallpaperStyle.SelLength = 0
    cmbWallpaperTimerInterval.SelLength = 0
   
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(Optional ByRef currFont As String, Optional ByRef currSize As Integer, Optional ByRef currWeight As Integer, Optional ByRef currStyle As Boolean, Optional ByRef currColour, Optional ByRef currItalics As Boolean, Optional ByRef currUnderline As Boolean, Optional ByRef fontResult As Boolean)

       
    ' variables declared
    Dim f As FormFontInfo
        
    'initialise the dimensioned variables
    'f =
   
   On Error GoTo displayFontSelector_Error
   If debugflg = 1 Then Debug.Print "%displayFontSelector"

    With f
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = DialogFont(f)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If f.Name = vbNullString Then f.Name = "times new roman"
    If f.Name = vbNullString Then Exit Sub
    
    With f
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of Form dockSettings"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(Index As Integer)
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuCoffee_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuCoffee_Click"
    
    answer = MsgBox(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelpPdf_click
' Author    : beededea
' Date      : 30/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpPdf_click()
       
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
   On Error GoTo mnuHelpPdf_click_Error
   If debugflg = 1 Then Debug.Print "%mnuHelpPdf_click"

    answer = MsgBox("This option opens a browser window and displays this tool's help. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        If fFExists(App.Path & "\help\SteamyDockSettings.html") Then
            Call ShellExecute(Me.hWnd, "Open", App.Path & "\help\SteamyDockSettings.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("The help file - SteamyDockSettings.html- is missing from the help folder.")
        End If
    End If

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpPdf_click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFacebook_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFacebook_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuFacebook_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuFacebook_Click"

    answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFacebook_Click of Form quartermaster"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuLatest_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLatest_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuLatest_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuLatest_Click"

    answer = MsgBox("Download latest version of the program - this button opens a browser window and connects to the widget download page where you can check and download the latest zipped file). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If


    On Error GoTo 0
    Exit Sub

mnuLatest_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLatest_Click of Form quartermaster"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicence_Click()
    On Error GoTo mnuLicence_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuLicence_Click"
        
    Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
    licence.Show

    On Error GoTo 0
    Exit Sub

mnuLicence_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_Click of Form quartermaster"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuSupport_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuSupport_Click"

    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSweets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSweets_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    

    On Error GoTo mnuSweets_Click_Error
       If debugflg = 1 Then Debug.Print "%" & "mnuSweets_Click"
    
    
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.amazon.co.uk/gp/registry/registry.html?ie=UTF8&id=A3OBFB6ZN4F7&type=wishlist", vbNullString, App.Path, 1)
    End If
    
    On Error GoTo 0
    Exit Sub

mnuSweets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSweets_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuWidgets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuWidgets_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo

    On Error GoTo mnuWidgets_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuWidgets_Click"
    
    answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuWidgets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuWidgets_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClose_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Close the program from the menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuClose_Click()
    On Error GoTo mnuClose_Click_Error
    If debugflg = 1 Then Debug.Print "mnuClose_Click"
    
    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Clickg_Click of Form dockSettings"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuDebug_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Run the runtime debugging window exectuable
'---------------------------------------------------------------------------------------
'
Private Sub mnuDebug_Click()
        
    On Error GoTo mnuDebug_Click_Error
    
    If debugflg = 1 Then Debug.Print "%mnuDebug_Click"
    
    If debugflg = 0 Then
        debugflg = 1
        mnuDebug.Caption = "Turn Developer Options OFF"
        mnuAppFolder.Visible = True
        mnuEditWidget.Visible = True
        
        lblGenLabel(5).Enabled = True
        lblGenLabel(6).Enabled = True
        lblGenLabel(0).Enabled = True
        lblGenLabel(1).Enabled = True
        
        txtDockDefaultEditor.Enabled = True
        txtDockSettingsDefaultEditor.Enabled = True
        txtIconSettingsDefaultEditor.Enabled = True
        
        btnGeneralDockEditor.Enabled = True
        btnGeneralDockSettingsEditor.Enabled = True
        btnGeneralIconSettingsEditor.Enabled = True
        
    Else
        debugflg = 0
        mnuDebug.Caption = "Turn Developer Options ON"
        mnuAppFolder.Visible = False
        mnuEditWidget.Visible = False
        
        lblGenLabel(5).Enabled = False
        lblGenLabel(6).Enabled = False
        lblGenLabel(0).Enabled = False
        lblGenLabel(1).Enabled = False
        
        txtDockDefaultEditor.Enabled = False
        txtDockSettingsDefaultEditor.Enabled = False
        txtIconSettingsDefaultEditor.Enabled = False
        
        btnGeneralDockEditor.Enabled = False
        btnGeneralDockSettingsEditor.Enabled = False
        btnGeneralIconSettingsEditor.Enabled = False
    End If

    gblRdDebugFlg = CStr(debugflg)
    PutINISetting "Software\SteamyDock\DockSettings", "debugFlg", gblRdDebugFlg, toolSettingsFile

   On Error GoTo 0
   Exit Sub

mnuDebug_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDebug_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click(Index As Integer)
    
    On Error GoTo mnuAbout_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuAbout_Click"
          
     about.lblMajorVersion.Caption = App.Major
     about.lblMinorVersion.Caption = App.Minor
     about.lblRevisionNum.Caption = App.Revision
     
     about.Show
     
     If (about.WindowState = 1) Then
         about.WindowState = 0
     End If


    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayVersionNumber
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub displayVersionNumber()
   On Error GoTo displayVersionNumber_Error
   If debugflg = 1 Then Debug.Print "%displayVersionNumber"

     dockSettings.lblMajorVersion.Caption = App.Major
     dockSettings.lblMinorVersion.Caption = App.Minor
     dockSettings.lblRevisionNum.Caption = App.Revision

   On Error GoTo 0
   Exit Sub

displayVersionNumber_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayVersionNumber of Form dockSettings"
End Sub



'An OLE_COLOR value is a BGR (Blue, Green, Red) value. To determine the BGR value, specify blue, green, or red (each of which has a value from 0 - 255) in the following formula:
'
'
'BGR Value = (blue * 65536) + (green * 256) + red
'
'
'r = 238
'G = 239
'B = 221
'
'
'The formula to convert to OLE_COLOR was:
'BGR Value = (blue * 65536) + (green * 256) + red
'
'
'
'221 * 65536 = 14483456 (Blue)
'239 * 256 = 61184          (Green)
'238                                (Red)
'
'
'14483456 + 61184 + 238 = 14544878 (Decimal) or &HDDEFEE (Hex)

'---------------------------------------------------------------------------------------
' Procedure : IsValidOleColor
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsValidOleColor(ByVal nColor As Long) As Boolean
   On Error GoTo IsValidOleColor_Error

  Select Case nColor
    Case 0& To &H100FFFF, &H2000000 To &H2FFFFFF
         IsValidOleColor = True
    Case &H80000000 To &H80FF0018
         IsValidOleColor = (nColor And &HFFFF&) <= &H18
  End Select

   On Error GoTo 0
   Exit Function

IsValidOleColor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsValidOleColor of Form dockSettings"
End Function



Private Sub sliStyleShadowOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleShadowOpacity.hWnd, "The strength of the shadow can be altered here.", _
                  TTIconInfo, "Help on the Shadow Opacity Slider.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleThemeSize_Change
' Author    : beededea
' Date      : 14/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleThemeSize_Change()

   On Error GoTo sliStyleThemeSize_Change_Error

    lblStyleSizeCurrent.Caption = "(" & Val(sliStyleThemeSize.Value) & "px)"
    
    rDSkinSize = sliStyleThemeSize.Value

   On Error GoTo 0
   Exit Sub

sliStyleThemeSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleThemeSize_Change of Form dockSettings"
End Sub



Private Sub sliStyleThemeSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleThemeSize.hWnd, "This controls the size of the background theme. Only implemented on SteamyDock.", _
                  TTIconInfo, "Help on Theme Size.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0

    ' This should only be required on a machine that can give the Windows classic theme to the UI
    ' that excludes windows 8 and 10 so this timer can be switched off on these o/s.

    On Error GoTo themeTimer_Timer_Error
   
    ' In the IDE the background sys colour is derived from the IDE and not from the program form so we disregard the discrepancy
    ' and avoid changing the background colour when running from within the IDE

    SysClr = GetSysColor(COLOR_BTNFACE)
    If debugflg = 1 Then Debug.Print "COLOR_BTNFACE = " & SysClr ' generates too many debug statements in the log
    
    If InIDE = False Then
        If debugflg = 1 Then debugLog "COLOR_BTNFACE = " & SysClr  ' generates too many debug statements in the log
        If SysClr <> storeThemeColour Then
            Call setThemeColour
        End If
    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Public Sub setThemeColour()
    
        
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0
    
   On Error GoTo setThemeColour_Error
   If debugflg = 1 Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        
        'load the gear images
        imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1.jpg")
        imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3.jpg")

        rDSkinTheme = "dark"
    Else
        'MsgBox "Windows Alternate Theme detected"
        SysClr = GetSysColor(COLOR_BTNFACE)
        If SysClr = 13160660 Then
            Call setThemeShade(212, 208, 199)
            rDSkinTheme = "dark"
            
                    
            'load the gear images
            imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1.jpg")
            imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1.jpg")

        Else ' 15790320
            Call setThemeShade(240, 240, 240)
            rDSkinTheme = "light"
        
            'load the gear images
            imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1Light.jpg")
            imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3Light.jpg")
            
        End If

    End If



    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setThemeSkin
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setThemeSkin()
   On Error GoTo setThemeSkin_Error

    If rDSkinTheme = "dark" Then
        'load the gear images
        imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1.jpg")
        imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3.jpg")
        Call setThemeShade(212, 208, 199)
    Else
        'load the gear images
        imgMultipleGears1.Picture = LoadPicture(App.Path & "\resources\images\multipleGears1Light.jpg")
        imgMultipleGears3.Picture = LoadPicture(App.Path & "\resources\images\multipleGears3Light.jpg")
        Call setThemeShade(240, 240, 240)
    End If

   On Error GoTo 0
   Exit Sub

setThemeSkin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeSkin of Form dockSettings"
End Sub





'    If fFExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        PutINISetting "Software\RocketDockSettings", "defaultFont", suppliedFont, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultSize", suppliedSize, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultStrength", suppliedStrength, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultStyle", suppliedStyle, toolSettingsFile
'    End If


'---------------------------------------------------------------------------------------
' Procedure : placeFrames
' Author    : beededea
' Date      : 09/05/2020
' Purpose   : place the frames for the icons and main tabs into the correct position and space
'---------------------------------------------------------------------------------------
'
Private Sub placeFrames()
        
    Dim top As Integer: top = 0
    Dim gap As Integer: gap = 0
    Dim useloop As Integer: useloop = 0
        
    On Error GoTo placeFrames_Error
   
    If debugflg = 1 Then Debug.Print "%placeFrames"
    
    top = 0
    gap = 1300
    useloop = 0
    
    ' icon frames

    fmeGeneral.top = top
    fmeIcons.top = fmeGeneral.top + gap
    fmeBehaviour.top = fmeIcons.top + gap
    fmeStyle.top = fmeBehaviour.top + gap + 125
    fmePosition.top = fmeStyle.top + gap
    fmeWallpaper.top = fmePosition.top + gap
    fmeAbout.top = fmeWallpaper.top + gap
    
    ' tab frames
    
    For useloop = 0 To 6
        fmeMain(useloop).Left = 1665
        fmeMain(useloop).top = 30
    Next useloop

   On Error GoTo 0
   Exit Sub

placeFrames_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure placeFrames of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : writeDockSettings
' Author    : beededea
' Date      : 12/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub writeDockSettings(location As String, settingsFile As String)

' Alternative settings.ini file called docksettings.ini
' partitioned as follows:
'
' [Software\SteamyDock\DockSettings]
' [Software\SteamyDock\IconSettings]
' [Software\SteamyDock\SteamyDock\DockSettings]

    On Error GoTo writeDockSettings_Error
    If debugflg = 1 Then Debug.Print "%writeDockSettings"
    
    ' first we save the Steamydock specific settings
    If fFExists(dockSettingsFile) Then
        PutINISetting "Software\SteamyDock\DockSettings", "GeneralReadConfig", rDGeneralReadConfig, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "GeneralWriteConfig", rDGeneralWriteConfig, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "RunAppInterval", rDRunAppInterval, dockSettingsFile
'        PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "DefaultDock", rDDefaultDock, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "AnimationInterval", rDAnimationInterval, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "SkinSize", rDSkinSize, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "SplashStatus", sDSplashStatus, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "ShowIconSettings", sDShowIconSettings, dockSettingsFile ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock

        PutINISetting "Software\SteamyDock\DockSettings", "FontOpacity", sDFontOpacity, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "AutoHideType", sDAutoHideType, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "ShowLblBacks", sDShowLblBacks, dockSettingsFile ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
        PutINISetting "Software\SteamyDock\DockSettings", "ContinuousHide", sDContinuousHide, dockSettingsFile   'nn
        PutINISetting "Software\SteamyDock\DockSettings", "BounceZone", sDBounceZone, dockSettingsFile   'nn
    'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "ContinuousHide", sDContinuousHide)
   ' ContinuousHide
   End If
    
    ' then we save those associated to both Rocketdock and SteamyDock
    PutINISetting location, "Version", rDVersion, settingsFile
    PutINISetting location, "HotKey-Toggle", rDHotKeyToggle, settingsFile
    PutINISetting location, "Theme", rDtheme, settingsFile
    PutINISetting location, "Wallpaper", rDWallpaper, settingsFile
    PutINISetting location, "WallpaperStyle", rDWallpaperStyle, settingsFile
    PutINISetting location, "AutomaticWallpaperChange", rDAutomaticWallpaperChange, settingsFile
    PutINISetting location, "WallpaperTimerIntervalIndex", rDWallpaperTimerIntervalIndex, settingsFile
    PutINISetting location, "WallpaperTimerInterval", rDWallpaperTimerInterval, settingsFile
    PutINISetting location, "MoveWinTaskbar", rDMoveWinTaskbar, settingsFile
    
    PutINISetting location, "ThemeOpacity", rDThemeOpacity, settingsFile
    PutINISetting location, "IconOpacity", rDIconOpacity, settingsFile
    PutINISetting location, "FontSize", rDFontSize, settingsFile
    PutINISetting location, "FontFlags", rDFontFlags, settingsFile
    PutINISetting location, "FontName", rDFontName, settingsFile
    PutINISetting location, "FontColor", rDFontColor, settingsFile
    PutINISetting location, "FontCharSet", rDFontCharSet, settingsFile
    PutINISetting location, "FontOutlineColor", rDFontOutlineColor, settingsFile
    PutINISetting location, "FontOutlineOpacity", rDFontOutlineOpacity, settingsFile
    PutINISetting location, "FontShadowColor", rDFontShadowColor, settingsFile
    PutINISetting location, "FontShadowOpacity", rDFontShadowOpacity, settingsFile
    PutINISetting location, "IconMin", rDIconMin, settingsFile
    PutINISetting location, "IconMax", rdIconMax, settingsFile
    PutINISetting location, "ZoomWidth", rDZoomWidth, settingsFile
    PutINISetting location, "ZoomTicks", rDZoomTicks, settingsFile
    PutINISetting location, "AutoHide", rDAutoHide, settingsFile '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
    PutINISetting location, "AutoHideTicks", rDAutoHideDuration, settingsFile
    PutINISetting location, "AutoHideDelay", rDAutoHideDelay, settingsFile
    PutINISetting location, "PopupDelay", rDPopupDelay, settingsFile
    PutINISetting location, "IconQuality", rDIconQuality, settingsFile
    PutINISetting location, "LangID", rDLangID, settingsFile
    PutINISetting location, "HideLabels", rDHideLabels, settingsFile
    PutINISetting location, "ZoomOpaque", rDZoomOpaque, settingsFile
    PutINISetting location, "LockIcons", rDLockIcons, settingsFile
    PutINISetting location, "RetainIcons", rDRetainIcons, settingsFile ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    
    PutINISetting location, "ManageWindows", rDManageWindows, settingsFile
    PutINISetting location, "DisableMinAnimation", rDDisableMinAnimation, settingsFile
    PutINISetting location, "ShowRunning", rDShowRunning, settingsFile
    PutINISetting location, "OpenRunning", rDOpenRunning, settingsFile
    PutINISetting location, "HoverFX", rDHoverFX, settingsFile
    PutINISetting location, "zOrderMode", rDzOrderMode, settingsFile
    PutINISetting location, "MouseActivate", rDMouseActivate, settingsFile
    PutINISetting location, "IconActivationFX", rDIconActivationFX, settingsFile
    PutINISetting location, "SoundSelection", rDSoundSelection, settingsFile
    
    PutINISetting location, "Monitor", rDMonitor, settingsFile
    PutINISetting location, "Side", rDSide, settingsFile
    PutINISetting location, "Offset", rDOffset, settingsFile
    PutINISetting location, "vOffset", rDvOffset, settingsFile
    PutINISetting location, "OptionsTabIndex", rDOptionsTabIndex, toolSettingsFile
    PutINISetting location & "\WindowFilters", "Count", 0, settingsFile
    
    ' this tool's local settings.ini
    PutINISetting "Software\DockSettings", "dockSettingsDefaultEditor", sDDockSettingsDefaultEditor, toolSettingsFile
        
    ' icon settings tool
    PutINISetting "Software\IconSettings", "iconSettingsDefaultEditor", gblSdIconSettingsDefaultEditor, iconSettingsToolFile
    
    ' the dock itself
    PutINISetting "Software\SteamyDock\DockSettings", "dockDefaultEditor", sDDockDefaultEditor, dockSettingsFile
       
   On Error GoTo 0
   Exit Sub

writeDockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeDockSettings of Form dockSettings"
End Sub






'---------------------------------------------------------------------------------------
' Procedure : adjustControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustControls()

        
    ' variables declared
    Dim rgbRdFontShadowColor As String
    Dim rgbRdFontOutlineColor As String

    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean

    
    'initialise the dimensioned variables
    rgbRdFontShadowColor = vbNullString
    rgbRdFontOutlineColor = vbNullString
    suppliedFontSize = 0
    suppliedWeight = 0
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False

    On Error GoTo adjustControls_Error
    If debugflg = 1 Then Debug.Print "%adjustControls"
    
    ' wallpaper controls
    
    Call populateWallpapers
    Call populateWallpaperStyleDropDown
    Call populateWallpaperTimerIntervalDropDown
    
    chkAutomaticWallpaperChange.Value = CInt(rDAutomaticWallpaperChange)
    cmbWallpaperTimerInterval.ListIndex = CInt(rDWallpaperTimerIntervalIndex)
    
    Call populateThemes

    optGeneralReadConfig.Value = CBool(rDGeneralReadConfig)
      If rDGeneralReadConfig = "True" Then
        optGeneralReadConfig.Value = True

      Else
        optGeneralReadConfig.Value = False

      End If
      

    'optGeneralWriteConfig.Value = CBool(LCase(rDGeneralWriteConfig))

      If rDGeneralWriteConfig = "True" Then
          optGeneralWriteConfig.Value = True
      Else
          optGeneralWriteConfig.Value = False
      End If


    ' controls for values that do not appear in Rocketdock
    If defaultDock = 1 Then
        sliRunAppInterval.Value = Val(rDRunAppInterval)
'        chkGenAlwaysAsk.Value = Val(rDAlwaysAsk)
    End If

    'Rocketdock values also used by Steamydock
    
    ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    chkLockIcons.Value = Val(rDLockIcons)
    chkRetainIcons.Value = Val(rDRetainIcons)

    chkOpenRunning.Value = Val(rDOpenRunning)
    chkShowRunning.Value = Val(rDShowRunning)
    chkGenMin.Value = Val(rDManageWindows)
    chkGenDisableAnim.Value = Val(rDDisableMinAnimation)

    If chkGenMin.Value = 0 Then
        chkGenDisableAnim.Enabled = False
    Else
        chkGenDisableAnim.Enabled = True
    End If
    
    If chkShowRunning.Value = 0 Then
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else
'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
    ' Icons tab
    
    Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
    Call setBounceTypes
    Call populateSoundSelectionDropDown

    
    chkMoveWinTaskbar.Value = CInt(rDMoveWinTaskbar)
    
    cmbIconsQuality.ListIndex = Val(rDIconQuality)
    sliIconsOpacity.Value = Val(rDIconOpacity)
    chkIconsZoomOpaque.Value = Val(rDZoomOpaque)
    sliIconsSize.Value = Val(rDIconMin)
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)
    
    sliIconsZoom.Value = Val(rdIconMax)
    
    'Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation

    sliIconsZoomWidth.Value = Val(rDZoomWidth)
    sliIconsDuration.Value = Val(rDZoomTicks)
    
    ' position

    cmbPositionMonitor.ListIndex = Val(rDMonitor)
    cmbPositionScreen.ListIndex = Val(rDSide)
    cmbPositionLayering.ListIndex = Val(rDzOrderMode)
    sliPositionCentre.Value = Val(rDOffset)
    sliPositionEdgeOffset.Value = Val(rDvOffset)
    
    'style panel
    
    sliStyleOpacity.Value = Val(rDThemeOpacity)
    chkStyleDisable.Value = Val(rDHideLabels)
    
    chkLabelBackgrounds.Value = Val(sDShowLblBacks) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files

    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Val(Abs(rDFontSize)) & "pt"

    'the colour data that comes from the registry is RGB decimal
    rgbRdFontShadowColor = Convert_Dec2RGB(rDFontShadowColor)
    lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & rgbRdFontShadowColor
    lblStyleFontOutlineTest.ForeColor = rDFontShadowColor

    rgbRdFontOutlineColor = Convert_Dec2RGB(rDFontOutlineColor)
    lblStyleOutlineColourDesc.Caption = "Outline Colour: " & rgbRdFontOutlineColor
    lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor

    lblPreviewFont.ForeColor = rDFontColor

    sliStyleFontOpacity.Value = Val(sDFontOpacity)
    sliStyleOutlineOpacity.Value = Val(rDFontOutlineOpacity)
    sliStyleShadowOpacity.Value = Val(rDFontShadowOpacity)
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)

    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)

    ' behaviour
    
    chkAutoHide.Value = Val(rDAutoHide)
' 226/10/2020 docksettings .05 DAEB  added a manual click to the autohide toggle checkbox
' a checkbox value assignment does not trigger a checkbox click for this checkbox (in a frame) as normally occurs and there is no equivalent 'change event' for a checkbox
' so to force it to trigger we need a call to the click event
    Call chkAutoHide_Click
    
    sliAutoHideDuration.Value = Val(rDAutoHideDuration)
    
    sliContinuousHide.Value = Val(sDContinuousHide) ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    
    'sDBounceZone
    
    sliAnimationInterval.Value = Val(rDAnimationInterval)
    sliStyleThemeSize.Value = Val(rDSkinSize)
    chkSplashStatus.Value = Val(sDSplashStatus)
    
    chkShowIconSettings.Value = Val(sDShowIconSettings) ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock

    
    sliBehaviourAutoHideDelay.Value = Val(rDAutoHideDelay)
    
    cmbAutoHideType.ListIndex = Val(sDAutoHideType)

    chkBehaviourMouseActivate.Value = Val(rDMouseActivate)
    sliBehaviourPopUpDelay.Value = Val(rDPopupDelay)
    
    ' if not then add the key combinations that are allowed for Steamydock
    
    Call setHidingKey
    
    ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
    If defaultDock = 1 Then
        imgThemeSample.Enabled = True
        lblStyleLabel(2).Enabled = True
        
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
    Else
        imgThemeSample.Enabled = False
        lblStyleLabel(2).Enabled = False

        sliStyleThemeSize.Enabled = False
        lblThemeSizeTextHigh.Enabled = False
        lblStyleSizeCurrent.Enabled = False
    End If
    
    ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
    
    
   On Error GoTo 0
   Exit Sub

adjustControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustControls of Form dockSettings on line " & Erl

End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : populateThemes
' Author    : beededea
' Date      : 07/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateThemes()
     
    'Dim MyFile As String
    Dim MyPath  As String
    Dim themePresent As Boolean
    Dim myName As String
    
   On Error GoTo populateThemes_Error

    MyPath = dockAppPath & "\Skins\" '"E:\Program Files (x86)\RocketDock\Skins\"
    themePresent = False

    If Not fDirExists(MyPath) Then
        MsgBox "WARNING - The skins folder is not present in the correct location " & rdAppPath
    End If
    
    myName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(MyPath & myName) And vbDirectory) = vbDirectory Then
             'Debug.Print MyName   ' Display entry only if it
          End If   ' it represents a directory.
       End If
       myName = Dir   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> vbNullString Then
        cmbStyleTheme.AddItem myName
        'MsgBox MyName
        Debug.Print myName   ' Display entry only if it
        If myName = rDtheme Then themePresent = True
       End If
    Loop

    ' if the theme is not in the list then make it none to ensure no corruption *1
    If themePresent = False Then rDtheme = "blank"

    If rDtheme = "Program Files" Or rDtheme = vbNullString Then
        cmbStyleTheme.Text = "blank"
    Else
        cmbStyleTheme.Text = rDtheme
    End If

   On Error GoTo 0
   Exit Sub

populateThemes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateThemes of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : populateWallpapers
' Author    : beededea
' Date      : 07/04/2025
' Purpose   : read the wallpaper folder and extract all image names to a combobox list, must be jpgs.
'---------------------------------------------------------------------------------------
'
Private Sub populateWallpapers()
     
    Dim MyPath  As String: MyPath = vbNullString
    Dim match   As String: match = vbNullString
    Dim wallpaperPresent As Boolean: wallpaperPresent = False
    Dim myName  As String: myName = vbNullString
    
    On Error GoTo populateWallpapers_Error

    MyPath = dockAppPath & "\Wallpapers\"
    wallpaperPresent = False

    If Not fDirExists(MyPath) Then
        MsgBox "WARNING - The Wallpapers folder is not present in the correct location " & App.Path
    End If
    
    cmbWallpaper.AddItem "none selected"
    
    myName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(MyPath & myName) And vbDirectory) = vbDirectory Then
             'Debug.Print MyName   ' Display entry only if it
          End If   ' it represents a directory.
       End If
       myName = Dir   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> vbNullString Then
            match = LCase$(Right$(myName, 4))
            If match = ".jpg" Or match = ".jpeg" Then
                cmbWallpaper.AddItem myName
                'Debug.Print myName   ' Display entry only if it
                If myName = rDWallpaper Then wallpaperPresent = True
            End If
       End If
    Loop

    ' if the wallpaper is not in the list then make it none to ensure no corruption *1
    If wallpaperPresent = False Then rDWallpaper = "none selected"

    If rDWallpaper = "Program Files" Or rDWallpaper = vbNullString Then
        cmbWallpaper.Text = "none selected"
    Else
        cmbWallpaper.Text = rDWallpaper
    End If

   On Error GoTo 0
   Exit Sub

populateWallpapers_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateWallpapers of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setBounceTypes
' Author    : beededea
' Date      : 31/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setBounceTypes()

'None
'UberIcon Effects
'Bounce

   On Error GoTo setBounceTypes_Error

    cmbIconActivationFX.Clear

'    If defaultDock = 0 Then
'        cmbIconActivationFX.AddItem "None", 0
'        cmbIconActivationFX.AddItem "UberIcon Effects", 1
'        cmbIconActivationFX.AddItem "Bounce", 2
'        'rDIconActivationFX = "2"
'
'    Else
        cmbIconActivationFX.AddItem "None", 0
        cmbIconActivationFX.AddItem "Bounce", 1
        cmbIconActivationFX.AddItem "Miserable", 2
        'rDIconActivationFX = "1"
'    End If
    
    cmbIconActivationFX.ListIndex = Val(rDIconActivationFX)
    

   On Error GoTo 0
   Exit Sub

setBounceTypes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setBounceTypes of Form dockSettings"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : populateSoundSelectionDropDown
' Author    : beededea
' Date      : 31/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateSoundSelectionDropDown()

   On Error GoTo populateSoundSelectionDropDown_Error

    cmbBehaviourSoundSelection.Clear

    cmbBehaviourSoundSelection.AddItem "None", 0
    cmbBehaviourSoundSelection.AddItem "Ting", 1
    cmbBehaviourSoundSelection.AddItem "Click", 2
    
    cmbBehaviourSoundSelection.ListIndex = Val(rDSoundSelection)
    

   On Error GoTo 0
   Exit Sub

populateSoundSelectionDropDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateSoundSelectionDropDown of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : populateWallpaperStyleDropDown
' Author    : beededea
' Date      : 31/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateWallpaperStyleDropDown()

    
    On Error GoTo populateWallpaperStyleDropDown_Error

    cmbWallpaperStyle.Clear

    cmbWallpaperStyle.AddItem "Centre", 0
    cmbWallpaperStyle.AddItem "Tile", 1
    cmbWallpaperStyle.AddItem "Stretch", 2
        
   On Error GoTo 0
   Exit Sub

populateWallpaperStyleDropDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateWallpaperStyleDropDown of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : populateWallpaperTimerIntervalDropDown
' Author    : beededea
' Date      : 31/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateWallpaperTimerIntervalDropDown()
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo populateWallpaperTimerIntervalDropDown_Error

    cmbWallpaperTimerInterval.Clear

    cmbWallpaperTimerInterval.AddItem "5 mins", 0
    cmbWallpaperTimerInterval.ItemData(0) = 5
    cmbWallpaperTimerInterval.AddItem "10 mins", 1
    cmbWallpaperTimerInterval.ItemData(1) = 10
    cmbWallpaperTimerInterval.AddItem "15 mins", 2
    cmbWallpaperTimerInterval.ItemData(2) = 15
    cmbWallpaperTimerInterval.AddItem "30 mins", 3
    cmbWallpaperTimerInterval.ItemData(3) = 30
    cmbWallpaperTimerInterval.AddItem "60 mins", 4
    cmbWallpaperTimerInterval.ItemData(4) = 60
    cmbWallpaperTimerInterval.AddItem "2 hours", 5
    cmbWallpaperTimerInterval.ItemData(5) = 120
    cmbWallpaperTimerInterval.AddItem "4 hours", 6
    cmbWallpaperTimerInterval.ItemData(6) = 240
    cmbWallpaperTimerInterval.AddItem "8 hours", 7
    cmbWallpaperTimerInterval.ItemData(7) = 480
    cmbWallpaperTimerInterval.AddItem "16 hours", 8
    cmbWallpaperTimerInterval.ItemData(8) = 960
    cmbWallpaperTimerInterval.AddItem "24 hours", 9
    cmbWallpaperTimerInterval.ItemData(9) = 1440
    cmbWallpaperTimerInterval.AddItem "2 days", 10
    cmbWallpaperTimerInterval.ItemData(10) = 2880
    cmbWallpaperTimerInterval.AddItem "3 days", 11
    cmbWallpaperTimerInterval.ItemData(11) = 4320
    cmbWallpaperTimerInterval.AddItem "5 days", 12
    cmbWallpaperTimerInterval.ItemData(12) = 7200
    cmbWallpaperTimerInterval.AddItem "7 days", 13
    cmbWallpaperTimerInterval.ItemData(13) = 10080 ' number of minutes stored in itemData array
        
    'rDWallpaperTimerInterval = cmbWallpaperTimerInterval.List(CStr(rDWallpaperTimerIntervalIndex))
    rDWallpaperTimerInterval = cmbWallpaperTimerInterval.ItemData(rDWallpaperTimerIntervalIndex)
    
   On Error GoTo 0
   Exit Sub

populateWallpaperTimerIntervalDropDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateWallpaperTimerIntervalDropDown of Form dockSettings"

End Sub

' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
'---------------------------------------------------------------------------------------
' Procedure : setZoomTypes
' Author    : beededea
' Date      : 29/04/2021
' Purpose   : Set the default zoom types available to the type of dock selected
'---------------------------------------------------------------------------------------
'
Private Sub setZoomTypes()

    On Error GoTo setZoomTypes_Error
    
    cmbIconsHoverFX.Clear

'    If defaultDock = 0 Then
'        cmbIconsHoverFX.AddItem "None", 0
'        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
'        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
'        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
'        rDHoverFX = "1"
'
'    Else
        cmbIconsHoverFX.AddItem "None", 0
        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
        cmbIconsHoverFX.AddItem "Zoom: Bumpy", 4
'    End If
    
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)

    On Error GoTo 0
    Exit Sub

setZoomTypes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setZoomTypes of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetMonitorCount
' Author    : beededea
' Date      : 02/03/2020
' Purpose   : populate the monitor dropdown according to the number of monitors available
'---------------------------------------------------------------------------------------
'
Private Sub GetMonitorCount()
    
    ' variables declared
   Dim numberOfMonitors As Integer
   Dim useloop As Integer
    
   'initialise the dimensioned variables
   numberOfMonitors = 1
   useloop = 1
    
   On Error GoTo GetMonitorCount_Error
   If debugflg = 1 Then Debug.Print "%GetMonitorCount"

   numberOfMonitors = GetSystemMetrics(SM_CMONITORS)
   
   If numberOfMonitors <= 1 Then
        cmbPositionMonitor.Clear
        cmbPositionMonitor.AddItem "Monitor 1"
        cmbPositionMonitor.ListIndex = 0
        cmbPositionMonitor.Enabled = False
    Else
        'clear and populate the monitor list
        cmbPositionMonitor.Clear
        For useloop = 1 To numberOfMonitors
            cmbPositionMonitor.AddItem "Monitor " & useloop
        Next useloop
        cmbPositionMonitor.ListIndex = rDMonitor
   End If
   lblPositionMonitor.ToolTipText = "This computer has this many screens - " & numberOfMonitors
   cmbPositionMonitor.ToolTipText = "This computer has this many screens - " & numberOfMonitors

   On Error GoTo 0
   Exit Sub

GetMonitorCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetMonitorCount of Form dockSettings"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setToolTips
' Author    : beededea
' Date      : 27/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setToolTips()
    
    Dim aboutText As String: aboutText = vbNullString
    Dim wallpaperText As String: wallpaperText = vbNullString

    On Error GoTo setToolTips_Error

    If chkToggleDialogs.Value = 0 Then
        Call DestroyToolTip ' destroys the current tooltip
        
        rDEnableBalloonTooltips = "0" ' this is the flag used to determine whether a new balloon tooltip is generated
        
        btnDefaults.ToolTipText = "Revert ALL settings to the defaults"
        chkToggleDialogs.ToolTipText = "When checked this toggle will display the information pop-ups and balloon tips"
        btnHelp.ToolTipText = "Click here to open tool's HTML help page in your browser"
        picBusy.ToolTipText = "The program is doing something..."
        btnClose.ToolTipText = "Exit this utility"
        btnApplyWallpaper.ToolTipText = "Display the selected wallpaper on the desktop"
        btnSaveRestart.ToolTipText = "This will save your changes and restart the dock."
        lblText(0).ToolTipText = "General Configuration Options"
        
        imgIcon(0).ToolTipText = "General Configuration Options"
        imgIcon(1).ToolTipText = "Icon size and quality"
        imgIcon(2).ToolTipText = "Icon bounce and pop up effects"
        imgIcon(3).ToolTipText = "Dock effects and quality"
        imgIcon(4).ToolTipText = "Dock theme and font configuration"
        imgIcon(5).ToolTipText = "Desktop Wallpaper settings"
        imgIcon(6).ToolTipText = "About this program"
        
        chkShowIconSettings.ToolTipText = "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
        chkSplashStatus.ToolTipText = "Show Splash Screen on Start-up"
        
        btnGeneralDockEditor.ToolTipText = "Select the VB6 project file to allow editing of the dock itself from the developer menu."
        btnGeneralDockSettingsEditor.ToolTipText = "Select the VB6 project file to allow editing of this dock settings utility using the developer menu."
        btnGeneralIconSettingsEditor.ToolTipText = "Select the VB6 project file to allow editing of the icon settings tool from its developer menu."
        
        optGeneralReadSettings.ToolTipText = "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
        optGeneralReadRegistry.ToolTipText = "Stores the configuration where Rocketdock stores it, in the Registry, increasingly incompatible with Windows new standards, causes some security problems and requires admin rights to operate."
        optGeneralReadConfig.ToolTipText = "This stores ALL configuration within the user data area retaining future compatibility in Windows. The trouble is, only SteamyDock can access it."
        
        sliRunAppInterval.ToolTipText = "The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppInterval2.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppInterval3.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppIntervalCur.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenLabel(0).ToolTipText = "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
'        chkGenAlwaysAsk.ToolTipText = "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as above"
        btnGeneralRdFolder.ToolTipText = "Select the folder location of Rocketdock here"
        chkShowRunning.ToolTipText = "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency below"
        chkGenDisableAnim.ToolTipText = "If you dislike the minimise animation, click this"
        chkOpenRunning.ToolTipText = "If you click on an icon that is already running then it can open it or fire up another instance"
        txtAppPath.ToolTipText = "This is the extrapolated location of the currently selected dock. This is for information only."
        'cmbDefaultDock.ToolTipText = "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
        chkLockIcons.ToolTipText = "This is an essential option that stops you accidentally deleting your dock icons, click it!"
        
        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        chkRetainIcons.ToolTipText = "Dragging a program binary to the dock can take an automatically selected icon or you can retain the embedded icon."
        
        chkGenMin.ToolTipText = "This allows running applications to appear in the dock"
        chkStartupRun.ToolTipText = "This will cause the current dock to run when Windows starts"
        
'        optGeneralWriteSettings.ToolTipText = "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
        'optGeneralWriteRegistry.ToolTipText = "Stores the configuration where Rocketdock stores it, in the Registry, increasingly incompatible with Windows new standards, causes some security problems and requires admin rights to operate."
        optGeneralWriteConfig.ToolTipText = "This stores ALL configuration within the user data area retaining future compatibility in Windows. The trouble is, only SteamyDock can access it."

        'lblChkSplashStartup.ToolTipText = "Show Splash Screen on Start-up"
        'lblChkAlwaysConfirm.ToolTipText = "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as above"
        'lblChkOpenRunning.ToolTipText = "If you click on an icon that is already running then it can open it or fire up another instance"
        'lblRdLocation.ToolTipText = "This is the extrapolated location of the RocketDock Program, you can alter it yourself  if you have another copy of Rocketdock installed elsewhere - currently not operational, defaults to Rocketdock"
        
        lblGenLabel(2).ToolTipText = "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
        cmbHidingKey.ToolTipText = "This is the key sequence that is used to hide or restore Steamydock"
        sliContinuousHide.ToolTipText = "Determine how long Steamydock will disappear when told to hide using F11"
        
        cmbIconActivationFX.ToolTipText = "Set which type of animation you want to occur on an icon mouseover. Note SteamyDock will NOT support the Ubericon effects where Rocketdock does."
        chkAutoHide.ToolTipText = "You can determine whether the dock will auto-hide or not"
        sliAutoHideDuration.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        sliBehaviourPopUpDelay.ToolTipText = "The dock mouse-over delay period"
        lblBehaviourPopUpDelayMsCurrrent.ToolTipText = "The dock mouse-over delay period"
        sliBehaviourAutoHideDelay.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        chkBehaviourMouseActivate.ToolTipText = "Essential functionality for the dock - pops up when  given focus"
        lblBehaviourLabel(0).ToolTipText = "which type of animation you want to occur on an icon mouseover. Note SteamyDock will NOT support the Ubericon effects but Rocketdock will."
        'lblBehaviourLabel(1).ToolTipText = "You can determine whether the dock will auto-hide or not"
        lblBehaviourLabel(2).ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblBehaviourLabel(3).ToolTipText = "The dock mouse-over delay period"
        lblBehaviourLabel(4).ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblBehaviourLabel(5).ToolTipText = "Determine how long Steamydock will disappear when told to hide for the next few minutes"
        lblBehaviourLabel(6).ToolTipText = "This is the key sequence that is used to hide or restore Steamydock"
        lblBehaviourLabel(7).ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblBehaviourLabel(8).ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblBehaviourLabel(9).ToolTipText = "The dock mouse-over delay period"
        lblBehaviourLabel(10).ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblBehaviourLabel(11).ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        lblBehaviourLabel(12).ToolTipText = "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
        lblBehaviourLabel(13).ToolTipText = "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
        lblBehaviourLabel(14).ToolTipText = "The original icons may be low quality."
        lblBehaviourLabel(15).ToolTipText = "Select a sound to play when an icon in the dock is clicked."
        
        cmbBehaviourSoundSelection.ToolTipText = "Select a sound to play when an icon in the dock is clicked."
        
        lblContinuousHideMsCurrent.ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        lblContinuousHideMsHigh.ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        fraAutoHideType.ToolTipText = "The type of auto-hide, fade, instant or a slide like Rocketdock"
        lblAutoHideDurationMsHigh.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblAutoHideDurationMsCurrent.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblAutoRevealDurationMsHigh.ToolTipText = "The dock mouse-over delay period"
        lblAutoHideDelayMsHigh.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblAutoHideDelayMsCurrent.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        sliAnimationInterval.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second. The optimal value is probably 10ms."
        lblAnimationIntervalMsLow.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblAnimationIntervalMsHigh.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblAnimationIntervalMsCurrent.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        btnDonate.ToolTipText = "Opens a browser window and sends you to our donate page on Amazon"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs used by Rocketdock"
        btnFacebook.ToolTipText = "This will link you to the Rocket/Steamy dock users Group"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        chkLabelBackgrounds.ToolTipText = "You can toggle the icon label background on/off here"
        imgThemeSample.ToolTipText = "An example preview of the chosen theme."
        sliStyleShadowOpacity.ToolTipText = "The strength of the shadow can be altered here"
        sliStyleOutlineOpacity.ToolTipText = "The label outline transparency, use the slider to change"
        sliStyleFontOpacity.ToolTipText = "The font transparency can be changed here"
        lblStyleFontOpacityCurrent.ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(3).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleOutlineOpacityCurrent.ToolTipText = "The label outline transparency, use the slider to change"
        Label35.ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleShadowOpacityCurrent.ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(9).ToolTipText = "The strength of the shadow can be altered here"
        picStylePreview.ToolTipText = "A preview of the font selection - you can change the background of the preview to approximate how your font will look  on your desktop"
        btnStyleOutline.ToolTipText = "The colour of the outline, click the button to change"
        btnStyleShadow.ToolTipText = "The colour of the shadow, click the button to change"
        btnStyleFont.ToolTipText = "The font used in the labels, click the button to change"
        chkStyleDisable.ToolTipText = "You can toggle the icon labels on/off here"
        cmbStyleTheme.ToolTipText = "The dock background theme can be selected here"
        sliStyleOpacity.ToolTipText = "The theme background opacity is here"
        sliStyleThemeSize.ToolTipText = "The theme background overall size is here"
        
        lblChkLabelBackgrounds.ToolTipText = "You can toggle the icon label background on/off here"
        lblStyleLabel(2).ToolTipText = "The theme background overall size is here"
        lblStyleFontFontShadowColor.ToolTipText = "The colour of the shadow, click the button to change"
        lblStyleFontOutlineTest.ToolTipText = "The colour of the outline, click the button to change"
        lblStyleFontFontShadowTest.ToolTipText = "The colour of the shadow, click the button to change"
        lblStyleFontName.ToolTipText = "The font used in the labels, click the button to change"
        
        lblStyleLabel(0).ToolTipText = "The dock background theme can be selected here"
        lblStyleLabel(1).ToolTipText = "The theme background opacity is set here"
        lblStyleLabel(2).ToolTipText = "The theme background overall size is set here"
        lblStyleLabel(3).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(6).ToolTipText = "The theme background opacity is set here"
        lblStyleLabel(7).ToolTipText = "The theme background overall size is set here"
        lblStyleLabel(8).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(9).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(10).ToolTipText = "The label outline transparency, use the slider to change"
        

        fmeMain(0).ToolTipText = "These are the main settings for the dock"
        fmeMain(1).ToolTipText = "This panel allows you to set the icon sizes and hover effects"
        fmeMain(2).ToolTipText = "Here you can control the behaviour of the animation effects"
        fmeMain(3).ToolTipText = "This panel allows you to change the styling of the icon labels and the dock background image"
        fmeMain(4).ToolTipText = "This panel controls the positioning of the whole dock"
        fmeMain(5).ToolTipText = "This panel allows you to select a desktop background"
        fmeMain(6).ToolTipText = "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
        cmbPositionLayering.ToolTipText = "Should the dock appear on top of other windows or underneath?"
        cmbPositionMonitor.ToolTipText = "Here you can determine upon which monitor the dock will appear"
        cmbPositionScreen.ToolTipText = "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
        sliPositionEdgeOffset.ToolTipText = "Position from the bottom/top edge of the screen"
        sliPositionCentre.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label33.ToolTipText = "Should the dock appear on top of other windows or underneath?"
        lblPositionMonitor.ToolTipText = "Here you can determine upon which monitor the dock will appear"
        Label32.ToolTipText = "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
        Label31.ToolTipText = "You can align the dock so that it is centred or offas you require"
        lblPositionCentrePercCurrent.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label29.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label28.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label27.ToolTipText = "Position from the bottom/top edge of the screen"
        lblPositionEdgeOffsetPxCurrent.ToolTipText = "Position from the bottom/top edge of the screen"
        Label25.ToolTipText = "Position from the bottom/top edge of the screen"
        Label24.ToolTipText = "Position from the bottom/top edge of the screen"
        picMinSize.ToolTipText = "The icon size in the dock when static"
        picZoomSize.ToolTipText = "The maximum icon size of an animated icon"
        Label1.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label9.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label13.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        sliIconsDuration.ToolTipText = "How long the effect is applied"
        sliIconsZoomWidth.ToolTipText = "How many icons to the left and right are also animated"
        lblCharacteristicsLabel(11).ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(12).ToolTipText = "How long the effect is applied"
        lblIconsDurationMsCurrent.ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(6).ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(10).ToolTipText = "How many icons to the left and right are also animated"
        Label14.ToolTipText = "How many icons to the left and right are also animated"
        lblIconsZoomWidth.ToolTipText = "How many icons to the left and right are also animated"
        lblCharacteristicsLabel(5).ToolTipText = "How many icons to the left and right are also animated"
        chkIconsZoomOpaque.ToolTipText = "Should the zoom be opaque too?"
        cmbIconsQuality.ToolTipText = "Lower power single/dual core machines will benefit from the lower quality setting but in reality, current machines can run with high quality enabled and suffer no degradation whatsoever."
        sliIconsZoom.ToolTipText = "The maximum icon size after a zoom"
        sliIconsSize.ToolTipText = "The size of each icon in the dock before any effect is applied"
        sliIconsOpacity.ToolTipText = "The icons in the dock can be made transparent here"
        cmbIconsHoverFX.ToolTipText = "The zoom effect to apply"
        lblCharacteristicsLabel(2).ToolTipText = "The zoom effect to apply"
        lblCharacteristicsLabel(0).ToolTipText = "Lower power machines will benefit from the lower quality setting"
        lblCharacteristicsLabel(1).ToolTipText = "The icons in the dock can be made transparent here"
        lblCharacteristicsLabel(3).ToolTipText = "The size of each icon in the dock before any effect is applied"
        lblIconsOpacity.ToolTipText = "The icons in the dock can be made transparent here"
        lblIconsSize.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label3.ToolTipText = "The icons in the dock can be made transparent here"
        lblCharacteristicsLabel(7).ToolTipText = "The icons in the dock can be made transparent here"
        Label5.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        lblCharacteristicsLabel(8).ToolTipText = "The size of all the icons in the dock before any effect is applied"
        lblCharacteristicsLabel(4).ToolTipText = "The maximum icon size after a zoom"
        lblIconsZoom.ToolTipText = "The maximum icon size after a zoom"
        lblIconsZoomSizeMax.ToolTipText = "The maximum icon size after a zoom"
        lblCharacteristicsLabel(9).ToolTipText = "The maximum icon size after a zoom"
        picHiddenPicture.ToolTipText = "The icon size in the dock"
        Label26.ToolTipText = "Show Splash Screen on Start-up"
    
        aboutText = "About this program"
        fmeAbout.ToolTipText = aboutText
        imgIcon(6).ToolTipText = aboutText
        lblText(5).ToolTipText = aboutText
        
        wallpaperText = "This Button will select the wallpaper pane."
        fmeWallpaper.ToolTipText = wallpaperText
        imgIcon(5).ToolTipText = wallpaperText
        lblText(6).ToolTipText = wallpaperText
        
        chkAutomaticWallpaperChange.ToolTipText = "This checkbox enables a timer in the dock that will change the desktop background on an interval you define"
        imgWallpaperPreview.ToolTipText = "This image box displays a resized version of a much larger wallpaper, press the change button to apply it to your desktop."
        chkMoveWinTaskbar.ToolTipText = "If this is enabled, Steamydock will move the Windows taskbar to the opposite side, top to bottom &&c when the two overlap."
        
        btnNextWallpaper.ToolTipText = "To select the Next wallpaper click this button."
        btnPreviousWallpaper.ToolTipText = "To select the previous wallpaper click this button."
    Else
    
        rDEnableBalloonTooltips = "1" ' this is the flag used to determine whether a new balloon tooltip is generated
        
        ' module level balloon tooltip variables for comboBoxes ONLY.
        gcmbBehaviourActivationFXBalloonTooltip = "Set which type of animation you want to occur on an icon click. Only two animations are currently available. The bounce effect is best."
        gcmbBehaviourAutoHideTypeBalloonTooltip = "The type of dock auto-hide, fade away, instant hide or a slide-away like Rocketdock."
        gcmbHidingKeyBalloonTooltip = "This is the key sequence that is used to hide or restore Steamydock. Choose a key sequence that will not conflict with other apps you are running."
        gcmbBehaviourSoundSelectionBalloonTooltip = "Select a sound to play when an icon in the dock is clicked."
        gcmbStyleThemeBalloonTooltip = "The dock background theme can be selected here. The themes roughly match those available in Rocketdock."
        gcmbPositionMonitorBalloonTooltip = "Here you can determine upon which monitor the dock will appear."
        gcmbPositionScreenBalloonTooltip = "Place the dock at your preferred location. Steamydock currently only supports top and bottom positions." & vbCrLf & vbCrLf & "WARNING: In windows 10 it is not easy to move the taskbar, Windows 11 is even harder as it does not even give you the choice. However, we can do it programatically by changing the registry setting and then restarting explorer.exe. When you press save and restart, you will notice some flickering of the background as this change is made. "
        gcmbPositionLayeringBalloonTooltip = "This determines whether the dock should appear on top of other windows or underneath."
        gcmbIconsQualityBalloonTooltip = "Technically, lower power single/dual core machines from the XP era will benefit from the lower quality setting but in reality, the fast machines we have these days can run with high quality enabled and suffer no degradation whatsoever."
        gcmbIconsHoverFXBalloonTooltip = "The zoom effect to apply. At the moment the only effect in operation is bubble. The other animations types still need to be coded."
        gcmbDefaultDockBalloonTooltip = "This control merely indicates that the default dock is SteamyDock. Older versions worked with both Rocketdock and SteamyDock."
        gcmbWallpaperBalloonTooltip = "Select the wallpaper image that you desire to appear on the Windows desktop."
        gcmbWallpaperStyleBalloonTooltip = "Select the wallpaper style, centred, tiled or stretched."
        gcmbWallpaperTimerIntervalBalloonTooltip = "Select the interval at which the wallpaper will automatically change."

        cmbIconActivationFX.ToolTipText = vbNullString
        cmbAutoHideType.ToolTipText = vbNullString
        cmbHidingKey.ToolTipText = vbNullString
        cmbBehaviourSoundSelection.ToolTipText = vbNullString
        cmbStyleTheme.ToolTipText = vbNullString
        cmbPositionMonitor.ToolTipText = vbNullString
        cmbPositionScreen.ToolTipText = vbNullString
        cmbPositionLayering.ToolTipText = vbNullString
        cmbIconsQuality.ToolTipText = vbNullString
        cmbIconsHoverFX.ToolTipText = vbNullString
        cmbDefaultDock.ToolTipText = vbNullString

        btnDefaults.ToolTipText = vbNullString
        chkToggleDialogs.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        picBusy.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        btnApplyWallpaper.ToolTipText = vbNullString
        btnSaveRestart.ToolTipText = vbNullString
        lblText(0).ToolTipText = vbNullString
        imgIcon(0).ToolTipText = vbNullString
        imgIcon(1).ToolTipText = vbNullString
        imgIcon(2).ToolTipText = vbNullString
        imgIcon(3).ToolTipText = vbNullString
        imgIcon(4).ToolTipText = vbNullString
        imgIcon(5).ToolTipText = vbNullString
        imgIcon(6).ToolTipText = vbNullString
        chkShowIconSettings.ToolTipText = vbNullString
        chkSplashStatus.ToolTipText = vbNullString
        
        
        btnGeneralDockEditor.ToolTipText = vbNullString
        btnGeneralDockSettingsEditor.ToolTipText = vbNullString
        btnGeneralIconSettingsEditor.ToolTipText = vbNullString
        optGeneralReadSettings.ToolTipText = vbNullString
        optGeneralReadRegistry.ToolTipText = vbNullString
        optGeneralReadConfig.ToolTipText = vbNullString
        
        sliRunAppInterval.ToolTipText = vbNullString
        lblGenRunAppInterval2.ToolTipText = vbNullString
        lblGenRunAppInterval3.ToolTipText = vbNullString
        lblGenRunAppIntervalCur.ToolTipText = vbNullString
        'lblGenLabel(0).ToolTipText = vbNullString
        btnGeneralRdFolder.ToolTipText = vbNullString
        chkShowRunning.ToolTipText = vbNullString
        chkGenDisableAnim.ToolTipText = vbNullString
        chkOpenRunning.ToolTipText = vbNullString
        txtAppPath.ToolTipText = vbNullString
        chkLockIcons.ToolTipText = vbNullString
        chkRetainIcons.ToolTipText = vbNullString         ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        chkGenMin.ToolTipText = vbNullString
        chkStartupRun.ToolTipText = vbNullString

        optGeneralWriteConfig.ToolTipText = vbNullString

        sliContinuousHide.ToolTipText = vbNullString
        lblContinuousHideMsCurrent.ToolTipText = vbNullString
        lblContinuousHideMsHigh.ToolTipText = vbNullString
        fraAutoHideType.ToolTipText = vbNullString
        chkAutoHide.ToolTipText = vbNullString

        sliAutoHideDuration.ToolTipText = vbNullString
        lblAutoHideDurationMsHigh.ToolTipText = vbNullString
        lblAutoHideDurationMsCurrent.ToolTipText = vbNullString
        lblBehaviourPopUpDelayMsCurrrent.ToolTipText = vbNullString
        lblAutoRevealDurationMsHigh.ToolTipText = vbNullString
        sliBehaviourAutoHideDelay.ToolTipText = vbNullString
        sliBehaviourPopUpDelay.ToolTipText = vbNullString
        
        lblAutoHideDelayMsHigh.ToolTipText = vbNullString
        lblAutoHideDelayMsCurrent.ToolTipText = vbNullString
        lblBehaviourLabel(4).ToolTipText = vbNullString
        chkBehaviourMouseActivate.ToolTipText = vbNullString
        sliAnimationInterval.ToolTipText = vbNullString
        lblAnimationIntervalMsLow.ToolTipText = vbNullString
        lblAnimationIntervalMsHigh.ToolTipText = vbNullString
        lblAnimationIntervalMsCurrent.ToolTipText = vbNullString
        
        lblBehaviourLabel(2).ToolTipText = vbNullString
        lblBehaviourLabel(11).ToolTipText = vbNullString
        lblBehaviourLabel(5).ToolTipText = vbNullString
        lblBehaviourLabel(8).ToolTipText = vbNullString
        lblBehaviourLabel(9).ToolTipText = vbNullString
        lblBehaviourLabel(3).ToolTipText = vbNullString
        lblBehaviourLabel(10).ToolTipText = vbNullString
        lblBehaviourLabel(7).ToolTipText = vbNullString
        lblBehaviourLabel(12).ToolTipText = vbNullString
        
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        chkLabelBackgrounds.ToolTipText = vbNullString
        imgThemeSample.ToolTipText = vbNullString
        sliStyleShadowOpacity.ToolTipText = vbNullString
        sliStyleOutlineOpacity.ToolTipText = vbNullString
        sliStyleFontOpacity.ToolTipText = vbNullString

        lblStyleFontOpacityCurrent.ToolTipText = vbNullString
        lblStyleOutlineOpacityCurrent.ToolTipText = vbNullString
        lblStyleShadowOpacityCurrent.ToolTipText = vbNullString
        
        lblStyleLabel(0).ToolTipText = vbNullString
        lblStyleLabel(1).ToolTipText = vbNullString
        lblStyleLabel(2).ToolTipText = vbNullString
        lblStyleLabel(3).ToolTipText = vbNullString
        lblStyleLabel(4).ToolTipText = vbNullString
        lblStyleLabel(5).ToolTipText = vbNullString
        lblStyleLabel(6).ToolTipText = vbNullString
        lblStyleLabel(7).ToolTipText = vbNullString
        lblStyleLabel(8).ToolTipText = vbNullString
        lblStyleLabel(9).ToolTipText = vbNullString

        lblBehaviourLabel(13).ToolTipText = vbNullString
        lblBehaviourLabel(14).ToolTipText = vbNullString
        lblBehaviourLabel(15).ToolTipText = vbNullString
        
        cmbBehaviourSoundSelection.ToolTipText = vbNullString
        
        picStylePreview.ToolTipText = vbNullString
        btnStyleOutline.ToolTipText = vbNullString
        btnStyleShadow.ToolTipText = vbNullString
        btnStyleFont.ToolTipText = vbNullString
        chkStyleDisable.ToolTipText = vbNullString
        cmbStyleTheme.ToolTipText = vbNullString
        sliStyleOpacity.ToolTipText = vbNullString
        sliStyleThemeSize.ToolTipText = vbNullString
        
        lblChkLabelBackgrounds.ToolTipText = vbNullString
        lblStyleFontFontShadowColor.ToolTipText = vbNullString
        lblStyleFontOutlineTest.ToolTipText = vbNullString
        lblStyleFontFontShadowTest.ToolTipText = vbNullString
        lblStyleFontName.ToolTipText = vbNullString
        fmeMain(0).ToolTipText = vbNullString
        fmeMain(1).ToolTipText = vbNullString
        fmeMain(2).ToolTipText = vbNullString
        fmeMain(3).ToolTipText = vbNullString
        fmeMain(4).ToolTipText = vbNullString
        fmeMain(5).ToolTipText = vbNullString
        fmeMain(6).ToolTipText = vbNullString
        cmbPositionLayering.ToolTipText = vbNullString
        cmbPositionMonitor.ToolTipText = vbNullString
        cmbPositionScreen.ToolTipText = vbNullString
        sliPositionEdgeOffset.ToolTipText = vbNullString
        sliPositionCentre.ToolTipText = vbNullString
        Label33.ToolTipText = vbNullString
        lblPositionMonitor.ToolTipText = vbNullString
        Label32.ToolTipText = vbNullString
        Label31.ToolTipText = vbNullString
        lblPositionCentrePercCurrent.ToolTipText = vbNullString
        Label29.ToolTipText = vbNullString
        Label28.ToolTipText = vbNullString
        Label27.ToolTipText = vbNullString
        lblPositionEdgeOffsetPxCurrent.ToolTipText = vbNullString
        Label25.ToolTipText = vbNullString
        Label24.ToolTipText = vbNullString
        picMinSize.ToolTipText = vbNullString
        picZoomSize.ToolTipText = vbNullString
        Label1.ToolTipText = vbNullString
        Label9.ToolTipText = vbNullString
        Label13.ToolTipText = vbNullString
        sliIconsDuration.ToolTipText = vbNullString
        sliIconsZoomWidth.ToolTipText = vbNullString
        lblIconsDurationMsCurrent.ToolTipText = vbNullString
        lblCharacteristicsLabel(10).ToolTipText = vbNullString
        Label14.ToolTipText = vbNullString
        lblIconsZoomWidth.ToolTipText = vbNullString
        chkIconsZoomOpaque.ToolTipText = vbNullString
        cmbIconsQuality.ToolTipText = vbNullString
        sliIconsZoom.ToolTipText = vbNullString
        sliIconsSize.ToolTipText = vbNullString
        sliIconsOpacity.ToolTipText = vbNullString
        cmbIconsHoverFX.ToolTipText = vbNullString
        lblIconsOpacity.ToolTipText = vbNullString
        lblIconsSize.ToolTipText = vbNullString
        Label3.ToolTipText = vbNullString
        Label5.ToolTipText = vbNullString
        lblIconsZoom.ToolTipText = vbNullString
        lblIconsZoomSizeMax.ToolTipText = vbNullString
        
        lblCharacteristicsLabel(0).ToolTipText = vbNullString
        lblCharacteristicsLabel(2).ToolTipText = vbNullString
        lblCharacteristicsLabel(3).ToolTipText = vbNullString
        lblCharacteristicsLabel(4).ToolTipText = vbNullString
        lblCharacteristicsLabel(5).ToolTipText = vbNullString
        lblCharacteristicsLabel(6).ToolTipText = vbNullString
        lblCharacteristicsLabel(7).ToolTipText = vbNullString
        lblCharacteristicsLabel(8).ToolTipText = vbNullString
        lblCharacteristicsLabel(9).ToolTipText = vbNullString
        lblCharacteristicsLabel(11).ToolTipText = vbNullString
        lblCharacteristicsLabel(12).ToolTipText = vbNullString
        
        picHiddenPicture.ToolTipText = vbNullString
        Label26.ToolTipText = vbNullString
    
        aboutText = vbNullString
        fmeAbout.ToolTipText = aboutText
        imgIcon(6).ToolTipText = aboutText
        lblText(6).ToolTipText = aboutText
            
        wallpaperText = vbNullString
        fmeWallpaper.ToolTipText = wallpaperText
        imgIcon(5).ToolTipText = wallpaperText
        lblText(5).ToolTipText = wallpaperText
        
        chkAutomaticWallpaperChange.ToolTipText = vbNullString
        imgWallpaperPreview.ToolTipText = vbNullString
        
        chkMoveWinTaskbar.ToolTipText = vbNullString
        
        btnNextWallpaper.ToolTipText = vbNullString
        btnPreviousWallpaper.ToolTipText = vbNullString
    End If

    On Error GoTo 0
    Exit Sub

setToolTips_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setToolTips of Form dockSettings"
            Resume Next
          End If
    End With
End Sub

' .21 DAEB 07/09/2022 docksettings moved hiding key definitions to own subroutine
'---------------------------------------------------------------------------------------
' Procedure : setHidingKey
' Author    : beededea
' Date      : 07/09/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setHidingKey()

    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
    On Error GoTo setHidingKey_Error

    If defaultDock = 1 Then
        cmbHidingKey.Locked = False
        cmbHidingKey.Clear
        cmbHidingKey.AddItem "F1"
        cmbHidingKey.AddItem "F2"
        cmbHidingKey.AddItem "F3"
        cmbHidingKey.AddItem "F4"
        cmbHidingKey.AddItem "F5"
        cmbHidingKey.AddItem "F6"
        cmbHidingKey.AddItem "F7"
        cmbHidingKey.AddItem "F8"
        cmbHidingKey.AddItem "F9"
        cmbHidingKey.AddItem "F10"
        cmbHidingKey.AddItem "F11"
        cmbHidingKey.AddItem "F12"
        cmbHidingKey.AddItem "Disabled"
        cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    Else
        cmbHidingKey.Locked = True
        cmbHidingKey.Clear
        cmbHidingKey.AddItem "Control+Alt+R"
        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    End If
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD ends
    

    On Error GoTo 0
    Exit Sub

setHidingKey_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingKey of Form dockSettings"
            Resume Next
          End If
    End With
    
End Sub



Private Sub txtAppPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtAppPath.hWnd, "This is the extrapolated location of the currently selected dock. This is for information only.", _
                  TTIconInfo, "Help on the Running Application Indicators.", , , , True
End Sub
Private Sub positionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    dockSettingsXPos = dockSettings.Left
    dockSettingsYPos = dockSettings.top
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\DockSettings", "IconConfigFormXPos", dockSettingsXPos, toolSettingsFile
    PutINISetting "Software\DockSettings", "IconConfigFormYPos", dockSettingsYPos, toolSettingsFile
End Sub

Private Sub mnuBringToCentre_click()

    dockSettings.top = Screen.Height / 2 - dockSettings.Height / 2
    dockSettings.Left = screenWidthTwips / 2 - dockSettings.Width / 2
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
    Dim wallpaperPic As String: wallpaperPic = vbNullString
   
    On Error GoTo adjustMainControls_Error
          
    If sDDockSettingsDefaultEditor <> vbNullString Then
        mnuEditWidget.Caption = "Edit Program using " & sDDockSettingsDefaultEditor
        txtDockDefaultEditor.Text = sDDockDefaultEditor ' main steamydock editor location
        txtDockSettingsDefaultEditor.Text = sDDockSettingsDefaultEditor
        txtIconSettingsDefaultEditor.Text = gblSdIconSettingsDefaultEditor ' iconsettings editor location
    End If
    
    If debugflg = 1 Then
        mnuDebug.Caption = "Turn Developer Options OFF"
        mnuAppFolder.Visible = True
        mnuEditWidget.Visible = True
        
        lblGenLabel(5).Enabled = True
        lblGenLabel(6).Enabled = True
        lblGenLabel(0).Enabled = True
        lblGenLabel(1).Enabled = True
        
        txtDockDefaultEditor.Enabled = True
        txtDockSettingsDefaultEditor.Enabled = True
        txtIconSettingsDefaultEditor.Enabled = True
        
        btnGeneralDockEditor.Enabled = True
        btnGeneralDockSettingsEditor.Enabled = True
        btnGeneralIconSettingsEditor.Enabled = True
    Else
        mnuDebug.Caption = "Turn Developer Options ON"
        mnuAppFolder.Visible = False
        mnuEditWidget.Visible = False
        
        lblGenLabel(5).Enabled = False
        lblGenLabel(6).Enabled = False
        lblGenLabel(0).Enabled = False
        lblGenLabel(1).Enabled = False
        
        txtDockDefaultEditor.Enabled = False
        txtDockSettingsDefaultEditor.Enabled = False
        txtIconSettingsDefaultEditor.Enabled = False
        
        btnGeneralDockEditor.Enabled = False
        btnGeneralDockSettingsEditor.Enabled = False
        btnGeneralIconSettingsEditor.Enabled = False
    End If
    
    Call selectStoredWallpaperStyle
    Call selectStoredWallpaper
    
   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain"

End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : selectStoredWallpaperStyle
' Author    : beededea
' Date      : 10/04/2025
' Purpose   : select any previously stored and save wallpaper style
'---------------------------------------------------------------------------------------
'
 Private Sub selectStoredWallpaperStyle()
    Dim useloop As Integer: useloop = 0
    Dim thisStyle As String: thisStyle = vbNullString
    
   On Error GoTo selectStoredWallpaperStyle_Error

    thisStyle = rDWallpaperStyle

    'Iterate through items.
    For useloop = 0 To cmbWallpaperStyle.ListCount - 1
        'Compare value.
        If cmbWallpaperStyle.List(useloop) = thisStyle Then
            'Select it and leave loop.
            cmbWallpaperStyle.ListIndex = useloop
            Exit For
        End If
    Next useloop
    
    ' disable the apply button if no wallpaper choice
    If cmbWallpaperStyle.List(cmbWallpaperStyle.ListIndex) <> "none selected" Then
        btnApplyWallpaper.Enabled = True
    Else
        btnApplyWallpaper.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

selectStoredWallpaperStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure selectStoredWallpaperStyle of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : selectStoredWallpaper
' Author    : beededea
' Date      : 10/04/2025
' Purpose   : select any previously stored and save wallpaper
'---------------------------------------------------------------------------------------
'
Private Sub selectStoredWallpaper()
    
    Dim useloop As Integer: useloop = 0
    Dim wallpaperPic As String: wallpaperPic = vbNullString
    
    On Error GoTo selectStoredWallpaper_Error
    
    wallpaperPic = sdAppPath & "\wallpapers\" & rDWallpaper
    
    If fFExists(wallpaperPic) Then
        imgWallpaperPreview.Picture = LoadPicture(wallpaperPic)
    End If
    
    'cmbWallpaper.List = cmbWallpaper.List(Val(rDWallpaperTimerIntervalIndex))
    
    'Iterate through items.
    For useloop = 0 To cmbWallpaper.ListCount - 1
        'Compare value.
        If cmbWallpaper.List(useloop) = rDWallpaper Then
            'Select it and leave loop.
            cmbWallpaper.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

selectStoredWallpaper_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure selectStoredWallpaper of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : addTargetProgram
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Private Function addTargetProgram(ByVal targetText As String)
    Dim iconPath As String: iconPath = vbNullString
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    
    Const x_MaxBuffer = 256
    
    'On Error GoTo addTargetProgram_Error
    If debugflg = 1 Then debugLog "%" & "addTargetProgram"
    
    'On Error GoTo l_err1
    'savLblTarget = txtTarget.Text
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not targetText = vbNullString Then
        If fFExists(targetText) Then
            ' extract the folder name from the string
            iconPath = getFolderNameFromPath(targetText)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(targetText) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = targetText 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 dockSettings.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
            End If
        End If
    Else
    ' .85 DAEB 06/06/2022 rDIConConfig.frm  Second app button should open in the program files folder
    If fDirExists("c:\program files") Then
            dialogInitDir = "c:\program files"
        End If
    End If
    
    If Not sDockletFile = vbNullString Then
        If fFExists(sDockletFile) Then
            ' extract the folder name from the string
            dllPath = getFolderNameFromPath(sDockletFile)
            ' set the default folder to the existing reference
            dialogInitDir = dllPath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(sDockletFile) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = sDockletFile 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
                dialogInitDir = rdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            End If
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target for this icon to call"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String$(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With

  Call getFileNameAndTitle(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

  addTargetProgram = retFileName

   On Error GoTo 0
   
   Exit Function

addTargetProgram_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addTargetProgram of Form dockSettings"
 
End Function



'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
' NOTE: most of the settings are read from the SteamyDock settings file
'---------------------------------------------------------------------------------------
'

Public Sub readSettingsFile() '(ByVal location As String, ByVal PzGSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    ' this tool's local settings.ini
    If fFExists(toolSettingsFile) Then sDDockSettingsDefaultEditor = GetINISetting("Software\DockSettings", "dockSettingsDefaultEditor", toolSettingsFile)
        
    ' icon settings tool
    If fFExists(iconSettingsToolFile) Then gblSdIconSettingsDefaultEditor = GetINISetting("Software\IconSettings", "iconSettingsDefaultEditor", iconSettingsToolFile)
        
    ' the dock itself
    If fFExists(dockSettingsFile) Then sDDockDefaultEditor = GetINISetting("Software\SteamyDock\DockSettings", "dockDefaultEditor", dockSettingsFile)
       
        ' write the default editor for the icon settings.ini directly
        'If gblSdIconSettingsDefaultEditor <> vbNullString Then PutINISetting "Software\DockSettings", "dockDefaultEditor", gblSdIconSettingsDefaultEditor, iconSettingsToolFile
        
    gblRdDebugFlg = GetINISetting("Software\DockSettings", "debugFlg", toolSettingsFile)
    debugflg = Val(gblRdDebugFlg)

    gblFormPrimaryHeightTwips = GetINISetting("Software\DockSettings", "formPrimaryHeightTwips", dockSettingsFile)

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : subClassControls
' Author    : beededea
' Date      : 16/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub subClassControls()
    
   On Error GoTo subClassControls_Error

    If InIDE Then
        MsgBox "NOTE: Running in IDE so Sub classing is disabled" & vbCrLf & "Balloon tooltips will not display on comboboxes."
    Else
    
        ' sub classing code to intercept messages to the form itself in order to capture WM_EXITSIZEMOVE messages that occur AFTER the form has been resized
        
        Call SubclassForm(dockSettings.hWnd, ObjPtr(dockSettings))
        
        ' sub classing code to intercept messages to the comboboxes frame to provide missing balloon tooltips functionality
        Call SubclassComboBox(cmbIconActivationFX.hWnd, ObjPtr(cmbIconActivationFX))
        Call SubclassComboBox(cmbAutoHideType.hWnd, ObjPtr(cmbAutoHideType))
        Call SubclassComboBox(cmbHidingKey.hWnd, ObjPtr(cmbHidingKey))
        Call SubclassComboBox(cmbBehaviourSoundSelection.hWnd, ObjPtr(cmbBehaviourSoundSelection))
        Call SubclassComboBox(cmbStyleTheme.hWnd, ObjPtr(cmbStyleTheme))
        Call SubclassComboBox(cmbPositionMonitor.hWnd, ObjPtr(cmbPositionMonitor))
        Call SubclassComboBox(cmbPositionScreen.hWnd, ObjPtr(cmbPositionScreen))
        Call SubclassComboBox(cmbPositionLayering.hWnd, ObjPtr(cmbPositionLayering))
        Call SubclassComboBox(cmbIconsQuality.hWnd, ObjPtr(cmbIconsQuality))
        Call SubclassComboBox(cmbIconsHoverFX.hWnd, ObjPtr(cmbIconsHoverFX))
        Call SubclassComboBox(cmbDefaultDock.hWnd, ObjPtr(cmbDefaultDock))
        Call SubclassComboBox(cmbWallpaper.hWnd, ObjPtr(cmbWallpaper))
        Call SubclassComboBox(cmbWallpaperStyle.hWnd, ObjPtr(cmbWallpaperStyle))
        Call SubclassComboBox(cmbWallpaperTimerInterval.hWnd, ObjPtr(cmbWallpaperTimerInterval))
        
    End If

   On Error GoTo 0
   Exit Sub

subClassControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subClassControls of Form dockSettings"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : MouseMoveOnComboText
' Author    : beededea
' Date      : 16/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MouseMoveOnComboText(sComboName As String)
    Dim sTitle As String
    Dim sText As String

   On Error GoTo MouseMoveOnComboText_Error

    Select Case sComboName
    Case "cmbIconActivationFX"
        sTitle = "Help on Window Mode Selection."
        sText = gcmbBehaviourActivationFXBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbIconActivationFX.hWnd), sText, , sTitle, , , , True
    Case "cmbAutoHideType"
        sTitle = "Help on Open Running Behaviour."
        sText = gcmbBehaviourAutoHideTypeBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbAutoHideType.hWnd), sText, , sTitle, , , , True
    Case "cmbHidingKey"
        sTitle = "Help on the Hiding Key Selection"
        sText = gcmbHidingKeyBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbHidingKey.hWnd), sText, , sTitle, , , , True
    Case "cmbBehaviourSoundSelection"
        sTitle = "Help on the Sound Selection"
        sText = gcmbBehaviourSoundSelectionBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbBehaviourSoundSelection.hWnd), sText, , sTitle, , , , True
    Case "cmbStyleTheme"
        sTitle = "Help on the Style Theme Selection"
        sText = gcmbStyleThemeBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbStyleTheme.hWnd), sText, , sTitle, , , , True
    Case "cmbPositionMonitor"
        sTitle = "Help on the Monitor Position"
        sText = gcmbPositionMonitorBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbPositionMonitor.hWnd), sText, , sTitle, , , , True
    Case "cmbPositionScreen"
        sTitle = "Help on the Screen Position"
        sText = gcmbPositionScreenBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbPositionScreen.hWnd), sText, , sTitle, , , , True
    Case "cmbPositionLayering"
        sTitle = "Help on Layering the dock in relation to other Windows programs"
        sText = gcmbPositionLayeringBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbPositionLayering.hWnd), sText, , sTitle, , , , True
    Case "cmbIconsQuality"
        sTitle = "Help on the quality of the icons"
        sText = gcmbIconsQualityBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbIconsQuality.hWnd), sText, , sTitle, , , , True
    Case "cmbIconsHoverFX"
        sTitle = "Help on the icon hover effects"
        sText = gcmbIconsHoverFXBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbIconsHoverFX.hWnd), sText, , sTitle, , , , True
    Case "cmbDefaultDock"
        sTitle = "Help on the Default Dock"
        sText = gcmbDefaultDockBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbDefaultDock.hWnd), sText, , sTitle, , , , True
    Case "cmbWallpaper"
        sTitle = "Help on the Wallpaper Selection"
        sText = gcmbWallpaperBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbWallpaper.hWnd), sText, , sTitle, , , , True
    Case "cmbWallpaperStyle"
        sTitle = "Help on the Wallpaper Style"
        sText = gcmbWallpaperStyleBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbWallpaperStyle.hWnd), sText, , sTitle, , , , True
    Case "cmbWallpaperTimerInterval"
        sTitle = "Help on the Wallpaper Change Interval"
        sText = gcmbWallpaperTimerIntervalBalloonTooltip
        If rDEnableBalloonTooltips = "1" Then CreateToolTip cboEditHwndFromHwnd(cmbWallpaperTimerInterval.hWnd), sText, , sTitle, , , , True
    End Select

   On Error GoTo 0
   Exit Sub

MouseMoveOnComboText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseMoveOnComboText of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : loadHighlightedImages
' Author    : beededea
' Date      : 29/03/2025
' Purpose   : load the highlighted images onto the pressed icons
'---------------------------------------------------------------------------------------
'
Private Sub loadHighlightedImages()
   On Error GoTo loadHighlightedImages_Error

    imgIconPressed(0).Picture = LoadPicture(App.Path & "\resources\images\generalHighlighted.jpg")
    imgIconPressed(1).Picture = LoadPicture(App.Path & "\resources\images\iconsHighlighted.jpg")
    imgIconPressed(2).Picture = LoadPicture(App.Path & "\resources\images\behaviourHighlighted.jpg")
    imgIconPressed(3).Picture = LoadPicture(App.Path & "\resources\images\styleHighlighted.jpg")
    imgIconPressed(4).Picture = LoadPicture(App.Path & "\resources\images\positionHighlighted.jpg")
    imgIconPressed(5).Picture = LoadPicture(App.Path & "\resources\images\wallpaperHighlighted.jpg")
    imgIconPressed(6).Picture = LoadPicture(App.Path & "\resources\images\aboutHighlighted.jpg")

   On Error GoTo 0
   Exit Sub

loadHighlightedImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHighlightedImages of Form dockSettings"
End Sub


 '---------------------------------------------------------------------------------------
' Procedure : setDPIAware
' Author    : beededea
' Date      : 28/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setDPIAware()
    Const S_OK = &H0&, E_INVALIDARG = &H80070057, E_ACCESSDENIED = &H80070005

   On Error GoTo setDPIAware_Error

    Select Case SetProcessDpiAwareness(Process_System_DPI_Aware)
        'Case S_OK:           MsgBox "The current process is set as dpi aware.", vbInformation
        Case E_INVALIDARG:   MsgBox "The value passed in is not valid.", vbCritical
        Case E_ACCESSDENIED: MsgBox "The DPI awareness is already set, either by calling this API " & _
                                    "previously or through the application (.exe) manifest.", vbCritical
    End Select

   On Error GoTo 0
   Exit Sub

setDPIAware_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setDPIAware of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setFormHeight
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set the height of the whole form, only the form height is needed to proportion the form
'             constrain to not higher than the screen size, a resize causes a form_resize event
'---------------------------------------------------------------------------------------

Private Sub setFormHeight()

    On Error GoTo setFormHeight_Error
    
    ' constrain the height/width ratio
    gblConstraintRatio = pvtCFormHeight / pvtCFormWidth
     
    ' flag to cause a form's elements to all resize according to the new size set below
    gblFormResizedInCode = True
    
    ' set the form height using variables ready to test form height, not yet implemented
'    If  gblCurrentFormHeight < gblPhysicalScreenHeightTwips Then
       dockSettings.Height = CLng(gblFormPrimaryHeightTwips)
'    Else
'        dockSettings.Height = CLng(gblFormPrimaryHeightTwips) - 1000
'    End If

   On Error GoTo 0
   Exit Sub

setFormHeight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormHeight of Form dockSettings"
End Sub



    


'---------------------------------------------------------------------------------------
' Procedure : writeFormHeight
' Author    : beededea
' Date      : 21/05/2025
' Purpose   : write the form height to the settings file, only the form height is needed to proportion the form
'---------------------------------------------------------------------------------------
'
Private Sub writeFormHeight()
   On Error GoTo writeFormHeight_Error
   
    ' write the form height using variables ready to test dual monitors, that bit not yet implemented
    'If prefsMonitorStruct.IsPrimary = True Then
        gblFormPrimaryHeightTwips = Trim$(CStr(dockSettings.Height))
        PutINISetting "Software\DockSettings", "formPrimaryHeightTwips", gblFormPrimaryHeightTwips, dockSettingsFile
'    Else
'        gblPrefsSecondaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
'        sPutINISetting "Software\SteampunkClockCalendar", "prefsSecondaryHeightTwips", gblPrefsSecondaryHeightTwips, gblSettingsFile
'    End If

   On Error GoTo 0
   Exit Sub

writeFormHeight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeFormHeight of Form dockSettings"

End Sub
