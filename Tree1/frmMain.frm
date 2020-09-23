VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Your ..."
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2565
   FillColor       =   &H80000014&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   60
      TabIndex        =   25
      Top             =   4020
      Width           =   2445
      Begin VB.CheckBox chkTree 
         Caption         =   "On Top"
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   120
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Change Background"
         Height          =   315
         Index           =   8
         Left            =   180
         TabIndex        =   15
         Top             =   930
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fast Paint "
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Top             =   390
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Sounds"
         Height          =   315
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   660
         Value           =   1  'Checked
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   60
      TabIndex        =   19
      Top             =   630
      Width           =   2445
      Begin MSComctlLib.Slider sliderWidthSize 
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1710
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "2"
         Top             =   690
         Width           =   495
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "25"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Broken Branches "
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "35"
         Top             =   1350
         Width           =   495
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "75"
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fixed Angel"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1350
         Width           =   1815
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fixed Size"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "30"
         Top             =   210
         Width           =   645
      End
      Begin MSComctlLib.Slider sliderLength 
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   2100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sliderWind 
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sliderWidthScale 
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2700
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label Label2 
         Caption         =   "Width Scale:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   29
         Top             =   2700
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Wind :"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   28
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Width  :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   27
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Leaf Level :"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   2100
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Branches per Step :"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   3
         Left            =   2250
         TabIndex        =   23
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2250
         TabIndex        =   22
         Top             =   1350
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   1
         Left            =   2250
         TabIndex        =   21
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Total Steps :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop !"
      Height          =   405
      Left            =   1560
      TabIndex        =   17
      Top             =   5370
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Do it !"
      Default         =   -1  'True
      Height          =   405
      Left            =   735
      TabIndex        =   16
      Top             =   5370
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tree"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   60
      TabIndex        =   18
      Top             =   30
      Width           =   2445
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'       Real Tree (Branches & Leaves)
'       Copyright (c) 2006 Mohammad Reza Khosravi ( Khosravi2500@yahoo.com )
'
'1.     If you use values over 2 for  "Branches per Step", you must use
'       higher values for "Broken Branches", it makes a beatiful wild Tree.
'       [e.g. with "Total Steps"=30 use following values :(BPS=3, BB=52),(BPS=4, BB=65)
'       (BPS=5, BB=72),(BPS=6, BB=77),(BPS=7, BB=81),(BPS=8, BB=83)... ]
'
'2.     If you want to deactivate "Broken Branches", you must use lower values
'       for "Total Steps"  ( ~ 15, depends on speed of your computer)
'
'3.     For a classic fractal Tree try:
'       Total Steps =15
'       Branches per Step=2
'       Fixed Size = 75 (activated)
'       Fixed Angel= 35 (activated)
'       Broken Branches must be deactivated
'
'4.     Updates:
'       August 17... added "Leaf level" & "Width".
'       August 21... added "Change Background" & a few small changes in source code.
'       August 26... added "Wind Direction" & "OnTop" & improved for using higher values of "Branches per Step".
'       August 30... added "Width Scale".
'
'5.     If you find this program interesting (or not), please let me know.
'

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Sub Form_Load()
    Randomize Timer
    cmdStop.Left = cmdOK.Left
End Sub


Private Sub txtInput_GotFocus(Index As Integer)
    txtInput(Index).SelStart = 0
    txtInput(Index).SelLength = Len(txtInput(Index).Text)
End Sub


Private Sub chkTree_Click(Index As Integer)
    If Index > 4 Then Exit Sub
    If chkTree(Index).Value = 1 Then
        txtInput(Index).Enabled = True
        txtInput(Index).BackColor = vbWindowBackground
    Else
        txtInput(Index).Enabled = False
        txtInput(Index).BackColor = vbButtonFace
    End If

End Sub


Private Sub cmdOK_Click()
    Static noRun As Boolean
    
    cmdStop.Visible = True
    cmdOK.Visible = False
    
    If Not noRun Then ' only for the first time of running
        noRun = True
        Me.Left = 0
        Me.Top = 0
        frmTree.Left = Me.Width
        frmTree.Top = 0
        frmTree.Width = Screen.Width - Me.Width
        frmTree.Height = Screen.Height - 500
    End If
    
    ' restore variables from textboxes, checkboxes & sliders
    totalSteps = Abs(Val(txtInput(0).Text) Mod 100)
    totalSteps = IIf(totalSteps < 5, 5, totalSteps): txtInput(0).Text = totalSteps
    
    branchPerStep = Abs(Val(txtInput(1).Text) Mod 10)
    branchPerStep = IIf(branchPerStep < 2, 2, branchPerStep): txtInput(1).Text = branchPerStep
    
    fixSize = IIf(chkTree(2).Value = 1, Abs(Val(txtInput(2).Text) Mod 100), 0)
    fixAngel = IIf(chkTree(3).Value = 1, Abs(Val(txtInput(3).Text) Mod 360), 0)
    brokenBranches = IIf(chkTree(4).Value = 1, Abs(Val(txtInput(4).Text) Mod 100), 0) ' when you use higher steps, activating broken Branches make your speed faster and also better real Tree
    
    leafLevel = sliderLength.Value
    widthSize = sliderWidthSize.Value
    widthScale = sliderWidthScale.Value
    windDirection = sliderWind.Value
    
    onTop = IIf(chkTree(5).Value = 1, 0, 1)
    fastPaint = IIf(chkTree(6).Value = 1, True, False) ' if you use this option,your drawing speed is several time higher but when you move another Form on your Tree Form , your tree maybe became clear.
    playSound = IIf(chkTree(7).Value = 1, True, False)
    changeBackground = IIf(chkTree(8).Value = 1, True, False)
    
    'It's ready
    goRender = True
    If playSound Then sndPlaySound App.Path & "\start.wav", 1
    frmTree.Show
    Me.ZOrder onTop ' you can simply choose that this Form be on the top of the Tree Form or not.
End Sub


Private Sub cmdStop_Click()
    If playSound Then sndPlaySound App.Path & "\cancel.wav", 1
    goOut = True
End Sub


' Bye bye
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

