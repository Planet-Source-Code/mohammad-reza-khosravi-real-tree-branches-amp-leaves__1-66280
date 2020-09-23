VERSION 5.00
Begin VB.Form frmTree 
   AutoRedraw      =   -1  'True
   Caption         =   "...Tree"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "frmTree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6705
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5820
      Index           =   3
      Left            =   4380
      Picture         =   "frmTree.frx":0442
      ScaleHeight     =   5760
      ScaleWidth      =   8010
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   8070
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8160
      Index           =   2
      Left            =   2880
      Picture         =   "frmTree.frx":5444
      ScaleHeight     =   8100
      ScaleWidth      =   11520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   11580
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6300
      Index           =   1
      Left            =   5100
      Picture         =   "frmTree.frx":833B
      ScaleHeight     =   6240
      ScaleWidth      =   9000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6180
      Index           =   0
      Left            =   3630
      Picture         =   "frmTree.frx":D646
      ScaleHeight     =   6120
      ScaleWidth      =   9000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   9060
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi = 3.141592
Dim myStep As Byte

Private Sub Form_Activate()
    ' when you move between Forms, this procedure became active, but I want to run it once, so I use goRender variable
    If goRender Then goRender = False: Call renderMe
End Sub


Private Sub renderMe()
    
    refreshMe

    Me.AutoRedraw = Not fastPaint ' disabling of this property make faster drawing, but your tree is not stable and when you move another Form on the Tree Form, the tree may be clear.
    goOut = False ' this variable is for manually stop
    myStep = 0    ' this is the first step of function so it start from 0
    nextBranch Me.Width / 2, Me.Height - Me.Height / 5, 180
    
'    If Not goOut Then SavePicture Me.Image, App.Path & "\Pic" & Format(Now, "yyyymmdd-hhmmss") & ".bmp"  ' you can use this code for saving picture on hard disk, but "Fast Paint" must be OFF
    
    frmMain.cmdOK.Visible = True
    frmMain.cmdStop.Visible = False
    
End Sub


' this procedure only clean Form and show a new picture as background
Private Sub refreshMe()
    Static newBackground As Byte
    
    Me.AutoRedraw = True
    Me.Cls
    If changeBackground Then newBackground = Int(Rnd(1) * 4)
    Me.PaintPicture picBack(newBackground).Picture, 0, 0, Me.Width, Me.Height, 0, 0, picBack(newBackground).Width, picBack(newBackground).Height
End Sub


' this is the  main function that several times refer to itself
Private Function nextBranch(ByVal startX As Integer, ByVal startY As Integer, ByVal myDegree As Integer) As Boolean
    On Error GoTo mustOut 'this is for overflow error controlling when size of a branche became very high in random states
    Dim j As Byte
    Dim mySize As Single
    Dim endX As Integer, endY As Integer, myWidth As Integer, degreeGrow As Integer
    
    If myStep >= totalSteps Or goOut Then Exit Function
    If brokenBranches > 0 And myStep > 2 Then If Rnd(1) * 100 < brokenBranches Then Exit Function ' this is for making broken branches
    DoEvents
    
    myStep = myStep + 1
    
    
    ' different width for branches from root to leaves. if you are beginner try  myWidth=2
    myWidth = (widthSize / 5) * (15 + totalSteps / 4) / (myStep ^ (widthScale / 10 + 0.5)) ' in new update , for "Width Scale" action I added [^ (widthScale / 10 + 0.5)]
    If myWidth < 1 Then myWidth = 1

    ' length of branch. if you are beginner try  mySize = 500 ( in one line only ,delete following 3 lines) and also with lower values for "Total Steps"
    mySize = ((totalSteps - myStep * leafLevel / 5) * IIf(leafLevel >= 5, leafLevel / 5, (leafLevel + 5) / 10)) / (1 + Abs(leafLevel - 5) / 15)
    mySize = IIf(mySize > 0, mySize, 1) * (Me.Height / totalSteps ^ 1.9) * IIf(fixSize > 0, fixSize / 80, Rnd(1) * 1.5 + 0.1)

    If myStep < 3 Then mySize = mySize * (2 - myStep / 3) ' I added this statement in update (2) because I want higher length for trunk of tree. you can simply remove it !
       
    ' [ * Pi / 180 ] is for changing degrees to radians
    endX = Sin(myDegree * Pi / 180) * mySize + startX
    endY = Cos(myDegree * Pi / 180) * mySize + startY
    
    ' this paint tree branch
    Me.DrawWidth = myWidth
    Me.Line (startX, startY)-(endX, endY), RGB(100, 255 * myStep / totalSteps, 50) ' I prefer different colors from root to leaves
    
    ' this calculate degree for next branch, I made a change [* 2/branchPerStep] in  update3
    degreeGrow = (IIf(fixAngel > 0, fixAngel, (Int(Rnd(1) * 120) - 60))) * 2 / branchPerStep
    
    ' this is the place that function run itself again
    For j = 1 To branchPerStep
        nextBranch endX, endY, (myDegree * (1 - (windDirection - 5) / 100) - degreeGrow / 2 + degreeGrow * (j - branchPerStep / 2)) Mod 360 ' for Wind action, I added this statement in last update [* (1 - (windDirection - 5) / 100)]
    Next
    
    
    'nextBranch = True
mustOut:
    myStep = myStep - 1
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    goOut = True
End Sub

