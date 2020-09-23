VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "(Shaped) Pressure Soft Body"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   892
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Add Saved"
      Height          =   735
      Left            =   12480
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Me"
      Height          =   735
      Left            =   11280
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.HScrollBar sPressure 
      Height          =   255
      Left            =   11280
      Max             =   200
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CheckBox CHdOsHAPE 
      Caption         =   "Shaped"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   2
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10560
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00404040&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   120
      Width           =   10575
   End
   Begin VB.Label lINSIDE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label LabPress 
      Caption         =   "Inflated"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Started    As Boolean

Private Sub CHdOsHAPE_Click()
    Dim I      As Long
    For I = 1 To UBound(BO)

        BO(I).ShapeDO = IIf(CHdOsHAPE.Value = Checked, True, False)

    Next

End Sub

Private Sub Command1_Click()
    Started = True

    Dim I      As Long
    ReDim Preserve BO(3)

    For I = 1 To UBound(BO)
        BO(I).KStiffness = 250
        BO(I).KDamping = 0.85
        BO(I).Material = 1
        BO(I).COLOR = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        BO(I).ShapeDO = True
    Next


    BO(1).Create_BALL PicW / 2, PicH * 0.8, 25, 12, ByMuscle


    For I = 2 To 3
            BO(I).Create_Rectangle PicW * 0.1 + (20 * 6) * (I - 2), PicH * 0.5, 15, 2, 8, ByMuscle
            
            
        'BO(I).Create_BALL PicW * 0.1 + (20 * 6) * (I - 2), PicH * 0.5, 30, 12, ByInternalSprings
         '     BO(I).Create_BOX PicW * 0.1 + (55) * (I - 2), PicH * 0.5, 40, 80, ByInternalSprings

        'BO(I).Create_BALL PicW * 0.1 + (70) * (I - 2), PicH * 0.5, 30, 12, ByMuscle

    Next

    InitMatertials


    LabPress = "InflatedGas : " & BO(1).InflatedGas


    'BO(2).SetAllPointMasses = PositiveInfinity
    'BO(2).SetPointMass(1) = PositiveInfinity



    PTime = Timer
    DT = 0
    Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
    BO(1).Me_SAVE "TEST"
    BO(1).Me_LOAD "TEST", 100, 100
End Sub

Private Sub Command3_Click()
    ADDBODY "test", 100, 100

End Sub

Private Sub Form_Load()

    Randomize Timer

    PicW = PIC.Width
    PicH = PIC.Height
    ReDim BO(4)
    ReDim CollisionList(0)

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D      As Single
    Dim Dmin   As Single
    Dim I      As Long

    If Started = True Then

        If PointToMove = 0 Then
            Dmin = 99999999999#
            For I = 1 To BO(1).NumP
                D = VectorDist(Vector(BO(1).GetPointPosX(I), BO(1).GetPointPosY(I)), Vector(X, Y))
                If D < Dmin Then Dmin = D: PointToMove = I
            Next
        Else
            'PIC_MouseMove Button, Shift, x, y
        End If

    End If

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DIST   As tVector
    Dim Length As Single
    Dim Force  As Single
    If Started Then

        If Button = 1 Then


            DIST.X = X - BO(1).GetPointPosX(PointToMove)
            DIST.Y = Y - BO(1).GetPointPosY(PointToMove)

            Length = Sqr(DIST.X * DIST.X + DIST.Y * DIST.Y)
            ''    get spring force (Hookes law of elasticity)
            Force = -1 * Length
            ''    apply force to point mypoints #Pointtomove
            BO(1).SetPointVelX(PointToMove) = BO(1).GetPointVelX(PointToMove) - Force * (DIST.X / Length)
            BO(1).SetPointVelY(PointToMove) = BO(1).GetPointVelY(PointToMove) - Force * (DIST.Y / Length)
            '  Stop

        End If



        lINSIDE = IIf(BO(1).Contains(X, Y), "Mouse Inside", "Mouse Outside")
        



    Else

        If Button = 1 Then

            If BO(1).NumP = 0 Then
                BO(1).AddPoint X, Y, 0.5
            Else
                If 10 < VectorDist(Vector(X, Y), Vector(BO(1).GetPointPosX(BO(1).NumP), BO(1).GetPointPosY(BO(1).NumP))) Then

                    BO(1).AddPoint X, Y, 0.5
                    PIC.PSet (X, Y), vbWhite

                End If
            End If

        End If
    End If

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PointToMove = 0

    If Not Started Then
        BO(1).KStiffness = 250
        BO(1).KDamping = 0.85
        BO(1).Material = 1
        BO(1).COLOR = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

        BO(1).ShapeDO = True

        BO(1).SetUpSprings ByMuscle, False
        BO(1).AddALLMuscles

        BO(1).AREA = BO(1).GetAREA
        BO(1).Perimeter = BO(1).GetPerimeter
        BO(1).InflatedGas = 1.5 * BO(1).AREA / BO(1).Perimeter

        BO(1).Me_SAVE "test"
    End If

End Sub

Private Sub sPressure_Change()
    BO(1).InflatedGas = sPressure
    LabPress = "Inflated Gas : " & BO(1).InflatedGas
End Sub

Private Sub sPressure_Scroll()
    BO(1).InflatedGas = sPressure
    LabPress = "Inflated Gas : " & BO(1).InflatedGas
End Sub

Private Sub Timer1_Timer()
    Dim I      As Long
    Dim J      As Long


    DT = Timer - PTime
    PTime = Timer
    '    If DT > 0.025 Then DT = 0.025: PTime = Timer
    If DT > 0.25 Then DT = 0.25: PTime = Timer
    DT = 0.025


    For I = 1 To UBound(BO)
        BO(I).Do_Gravity
        BO(I).DO_FORCES
        BO(I).Do_ScreenBoundaries
    Next

    'PIC.Cls
    BitBlt PIC.hDC, 0, 0, PIC.ScaleWidth, PIC.ScaleHeight, PIC.hDC, 0, 0, vbBlackness

    For I = UBound(BO) To 1 Step -1
        BO(I).Me_DRAW PIC.hDC
    Next

    For I = 1 To UBound(BO) - 1
        For J = I + 1 To UBound(BO)
            BodyCollide I, J, CollisionList
            BodyCollide J, I, CollisionList
        Next
    Next

    HandleCollisions


    PIC.Refresh






End Sub
