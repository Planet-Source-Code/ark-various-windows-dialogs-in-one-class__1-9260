VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "CdlgEx test"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ShowAbout"
      Height          =   375
      Index           =   13
      Left            =   2280
      TabIndex        =   19
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowPrinter"
      Height          =   375
      Index           =   12
      Left            =   2280
      TabIndex        =   15
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowFont"
      Height          =   375
      Index           =   11
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowColor"
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowHelp"
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowSave"
      Height          =   375
      Index           =   8
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowOpen"
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SelectFolder"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FormatFloppy"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SelectIcon"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ObjectProperty"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reboot"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShutDown"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblDir 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CdlgEx1 As New CdlgEx

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = &H3
Const DI_COMPAT = &H4
Const DI_DEFAULTSIZE = &H8

Private Sub Command1_Click(Index As Integer)
 Dim retPath As String
 Dim LargeIcon As Boolean, IconSize As Long
    Select Case Index
           Case 0
                CdlgEx1.ShowShutDown
           Case 1
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.DialogPrompt = "Pressing Yes will restart Windows" & vbCrLf
                CdlgEx1.flags = Restart_Reboot
                CdlgEx1.Action = cdlgRestart
           Case 2
                CdlgEx1.DialogTitle = "Run your programm"
                CdlgEx1.DialogPrompt = "Ready to start?"
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.flags = Run_NoBrowse + Run_NoLable
                CdlgEx1.ShowRun ' CdlgEx1.hIcon
           Case 3
                CdlgEx1.flags = ObjProp_File
                CdlgEx1.ShowObjectProp "c:\"
           Case 4
                CdlgEx1.ShowFormat
           Case 5
                picIcon.AutoRedraw = True
                picIcon.Cls
                CdlgEx1.IconSize = IconSizeLarge
                CdlgEx1.ShowIcon
                picIcon.Width = CdlgEx1.IconSize * Screen.TwipsPerPixelX
                picIcon.Height = CdlgEx1.IconSize * Screen.TwipsPerPixelY
                DrawIconEx picIcon.hDC, 0, 0, CdlgEx1.hIcon, CdlgEx1.IconSize, CdlgEx1.IconSize, 0, 0, DI_NORMAL
                picIcon.Refresh
                lblFile = CdlgEx1.FileName
           Case 6
'                CdlgEx1.CancelError = True
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.DialogPrompt = "Choose folder for ..."
                CdlgEx1.flags = Folder_INCLUDEFILES
                CdlgEx1.InitDir = "c:\"
'                CdlgEx1.InitDir = ""  'TopFolder will be DeskTop
                CdlgEx1.SelDir = "c:\windows"
                CdlgEx1.ShowFolder
                lblDir = CdlgEx1.InitDir
           Case 7
                CdlgEx1.CancelError = True
                CdlgEx1.Filter = "Text files|*.txt|All files|*.*"
                CdlgEx1.ShowOpen
                lblFile = CdlgEx1.FileName
           Case 8
                CdlgEx1.CancelError = True
                CdlgEx1.Filter = "Text files|*.txt|All files|*.*"
                CdlgEx1.ShowSave
                lblFile = CdlgEx1.FileName
           Case 9
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.HelpFile = "c:\windows\help\windows.hlp"
                CdlgEx1.HelpCommand = HelpContents
                CdlgEx1.HelpKey = 0
                CdlgEx1.ShowHelp
                CdlgEx1.HelpCommand = HelpQuit
                CdlgEx1.ShowHelp
           Case 10
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.ShowColor
                Me.BackColor = CdlgEx1.RGBResult
           Case 11
                CdlgEx1.hOwner = Me.hWnd
                CdlgEx1.ShowFont
                Command1(Index).FontName = CdlgEx1.FontName
                Command1(Index).FontSize = CdlgEx1.FontSize
                Command1(Index).FontBold = CdlgEx1.Bold
                Command1(Index).FontItalic = CdlgEx1.Italic
           Case 12
                'Have no printer, so can not experiment w/ flags
                CdlgEx1.ShowPrinter
           Case 13
                CdlgEx1.AppName = "CommonDialog Extention"
                CdlgEx1.DialogPrompt = "CopyLeft (C) Ark Inc."
                CdlgEx1.ShowAbout
           Case Else
    End Select
End Sub

Private Sub Form_Load()
   CdlgEx1.FileName = "c:\autoexec.bat"
   CdlgEx1.InitDir = "c:\"
   CdlgEx1.SelDir = "c:\windows"
   Label1 = "Selected File"
   lblFile = CdlgEx1.FileName
   Label2 = "Selected Folder"
   lblDir = CdlgEx1.InitDir
   Label3 = "Selected Icon"
End Sub
