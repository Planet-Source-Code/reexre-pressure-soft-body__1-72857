VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoftBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSoftBody
' Author    : reexre@gmail.com Roberto Mior
' Date      : 22/01/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact

Option Explicit

Public NumP    As Long
Public NumS    As Long
Public NumSS   As Long


Private Mypoints() As tPoint
Private MySprings() As tSpring
Private ShapeSprings() As tSpring


Private NMuscles As Long    'Number of Muscles
Private Muscle() As tMuscle    'Muscles Array


Private AABB   As tAABB

Public Material As Long

Public ShapeMode As tShapeMode
Public ShapeDO As Boolean


Public KStiffness As Single
Public KDamping As Single



'Public Const InflatedGas as single = 20 ' 25 ' 400 ' 380
Public InflatedGas As Single


Public AREA    As Single
Public CurrAREA As Single
Public Perimeter As Single

Private BodyCenter As tVector
Private CurrBodyCenter As tVector
Private SpringDir As tVector
Private CurrSpringDir As tVector


Public SgnAREA As Integer

Public COLOR   As Long


Public Property Get GetPointPosX(wp) As Single
    GetPointPosX = Mypoints(wp).POS.X
End Property
Public Property Get GetPointPosY(wp) As Single
    GetPointPosY = Mypoints(wp).POS.Y
End Property
Public Property Get GetPointVelX(wp) As Single
    GetPointVelX = Mypoints(wp).VEL.X
End Property
Public Property Get GetPointVelY(wp) As Single
    GetPointVelY = Mypoints(wp).VEL.Y
End Property
Public Property Get GetPointForceX(wp) As Single
    GetPointForceX = Mypoints(wp).F.X
End Property
Public Property Get GetPointForceY(wp) As Single
    GetPointForceY = Mypoints(wp).F.Y
End Property

Public Property Let SetPointVelX(wp, V)
    Mypoints(wp).VEL.X = V
End Property
Public Property Let SetPointVelY(wp, V)
    Mypoints(wp).VEL.Y = V
End Property
Public Property Let SetPointPosX(wp, V)
    Mypoints(wp).POS.X = V
End Property
Public Property Let SetPointPosY(wp, V)
    Mypoints(wp).POS.Y = V
End Property



Public Property Get GetPointCpyPosX(wp) As Single
    GetPointCpyPosX = Mypoints(wp).CpyPOS.X
End Property
Public Property Get GetPointCpyPosY(wp) As Single
    GetPointCpyPosY = Mypoints(wp).CpyPOS.Y
End Property
Public Property Get GetPointCpyVelX(wp) As Single
    GetPointCpyVelX = Mypoints(wp).CpyVEL.X
End Property
Public Property Get GetPointCpyVelY(wp) As Single
    GetPointCpyVelY = Mypoints(wp).CpyVEL.Y
End Property

Public Property Get GetPointMASS(wp) As Single
    GetPointMASS = Mypoints(wp).Mass
End Property

Public Property Let SetPointMass(wp, Mass)
    Mypoints(wp).Mass = Mass

End Property

Public Property Let SetAllPointMasses(Mass)
    Dim I      As Long

    For I = 1 To NumP
        Mypoints(I).Mass = Mass
    Next

End Property




Public Property Get GetAABBMinX() As Single
    GetAABBMinX = AABB.Min.X
End Property
Public Property Get GetAABBMinY() As Single
    GetAABBMinY = AABB.Min.Y
End Property
Public Property Get GetAABBMaxX() As Single
    GetAABBMaxX = AABB.Max.X
End Property
Public Property Get GetAABBMaxY() As Single
    GetAABBMaxY = AABB.Max.Y
End Property

Public Function Atan2(ByVal dX As Double, ByVal dY As Double) As Double
'This Should return Angle

    Dim theta  As Double

    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            theta = 0#
        ElseIf (dY > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(dY / dX)

        If (dX < 0) Then
            If (dY >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If

    Atan2 = theta
End Function






Public Function Distance(dX As Single, dY As Single) As Single
    Distance = Sqr(dX * dX + dY * dY)

End Function

Public Sub AddSpring(wS, P1, P2)
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single

    MySprings(wS).P1 = P1
    MySprings(wS).P2 = P2
    MySprings(wS).Length = VectorDist(Mypoints(P1).POS, Mypoints(P2).POS)

    With MySprings(wS)
        X1 = Mypoints(.P1).POS.X
        Y1 = Mypoints(.P1).POS.Y
        X2 = Mypoints(.P2).POS.X
        Y2 = Mypoints(.P2).POS.Y

        .Normal.X = (Y1 - Y2) / .Length
        .Normal.Y = -(X1 - X2) / .Length
    End With

End Sub
Public Sub AddShapeSpring(P1, P2)
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    NumSS = NumSS + 1
    ReDim Preserve ShapeSprings(NumSS)



    ShapeSprings(NumSS).P1 = P1
    ShapeSprings(NumSS).P2 = P2
    ShapeSprings(NumSS).Length = VectorDist(Mypoints(P1).POS, Mypoints(P2).POS)

    With ShapeSprings(NumSS)
        X1 = Mypoints(.P1).POS.X
        Y1 = Mypoints(.P1).POS.Y
        X2 = Mypoints(.P2).POS.X
        Y2 = Mypoints(.P2).POS.Y

        .Normal.X = (Y1 - Y2) / .Length
        .Normal.Y = -(X1 - X2) / .Length
    End With

End Sub
Public Function ADDMuscle(L1, L2, Strength)
'
'
''Public Function ADDMuscle(L1, L2, Strength, _
  '            Optional dPhase360 As Single = 0, Optional dAmp360 As Single = 20, Optional dSpeed As Single = 0) As Boolean
'    'dPhase360 is the Dynamic Phase in Degrees
'    'dAmp360 is the Dynamic Angle Amplitude in Degrees (from -Value to +Value)
'    'dSpeed is the Dynamic Speed
'
    NMuscles = NMuscles + 1
    ReDim Preserve Muscle(NMuscles)

    ADDMuscle = False

    With Muscle(NMuscles)
        .L1 = L1
        .L2 = L2
        .F = Strength

        If MySprings(L1).P1 = MySprings(L2).P1 Then
            .P0 = MySprings(L1).P1
            .P1 = MySprings(L1).P2
            .P2 = MySprings(L2).P2
            ADDMuscle = True
        End If

        If MySprings(L1).P1 = MySprings(L2).P2 Then
            .P0 = MySprings(L1).P1
            .P1 = MySprings(L1).P2
            .P2 = MySprings(L2).P1
            ADDMuscle = True
        End If

        If MySprings(L1).P2 = MySprings(L2).P1 Then
            .P0 = MySprings(L1).P2
            .P1 = MySprings(L1).P1
            .P2 = MySprings(L2).P2
            ADDMuscle = True
        End If

        If MySprings(L1).P2 = MySprings(L2).P2 Then
            .P0 = MySprings(L1).P2
            .P1 = MySprings(L1).P1
            .P2 = MySprings(L2).P1
            ADDMuscle = True
        End If
        'Stop

    End With
    '
    If ADDMuscle = False Then
        NMuscles = NMuscles - 1
        ReDim Preserve Muscle(NMuscles)
        Exit Function
    End If



    With Muscle(NMuscles)
        .MainA = Atan2(Mypoints(.P1).POS.X - Mypoints(.P0).POS.X, Mypoints(.P1).POS.Y - Mypoints(.P0).POS.Y) - _
                 Atan2(Mypoints(.P2).POS.X - Mypoints(.P0).POS.X, Mypoints(.P2).POS.Y - Mypoints(.P0).POS.Y)


        'If dSpeed <> 0 Then
        '    .IsDynamic = True
        '     .DynPhase = dPhase360
        '     .DynAmp = dAmp360
        '     .DynSpeed = dSpeed
        ' End If

    End With

End Function


Public Sub AddPoint(XX, YY, Optional Mass = 0.5)
    NumP = NumP + 1
    ReDim Preserve Mypoints(NumP)
    Mypoints(NumP).POS.X = XX
    Mypoints(NumP).POS.Y = YY
    Mypoints(NumP).Mass = Mass

End Sub
Private Function GetBodyCenter(A As Single) As tVector
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    Dim I      As Long
    Dim V      As Single

    GetBodyCenter.X = 0
    GetBodyCenter.Y = 0

    For I = 1 To NumS
        With MySprings(I)
            X1 = Mypoints(.P1).POS.X
            Y1 = Mypoints(.P1).POS.Y
            X2 = Mypoints(.P2).POS.X
            Y2 = Mypoints(.P2).POS.Y


            '    V = (X1 * Y2 - X2 * Y1)
            '    GetBodyCenter.x = GetBodyCenter.x + (X1 + X2) * V
            '    GetBodyCenter.y = GetBodyCenter.y + (Y1 + Y2) * V

            GetBodyCenter.X = GetBodyCenter.X + (X1 + X2) * 0.5
            GetBodyCenter.Y = GetBodyCenter.Y + (Y1 + Y2) * 0.5

        End With
    Next


    'GetBodyCenter.x = GetBodyCenter.x / (6 * A)
    'GetBodyCenter.y = GetBodyCenter.y / (6 * A)

    GetBodyCenter.X = GetBodyCenter.X / NumS
    GetBodyCenter.Y = GetBodyCenter.Y / NumS

End Function

Public Function GetAREA() As Single
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    Dim I      As Long

    GetAREA = 0
    For I = 1 To NumS
        With MySprings(I)
            X1 = Mypoints(.P1).POS.X
            Y1 = Mypoints(.P1).POS.Y
            X2 = Mypoints(.P2).POS.X
            Y2 = Mypoints(.P2).POS.Y
            ''  MySprings(I).Normal.x = (Y1 - Y2) / .Length
            ''  MySprings(I).Normal.y = -(X1 - X2) / .Length
            ''  getarea = getarea + Abs(X1 - X2) * Abs(MySprings(I).Normal.x) * .Length
            'X1 = Mypoints(I).POS.x
            'Y1 = Mypoints(I).POS.y
            'X2 = Mypoints(I + 1).POS.x
            'Y2 = Mypoints(I + 1).POS.y

            GetAREA = GetAREA + (X1 * Y2 - X2 * Y1)
        End With

    Next


    GetAREA = GetAREA * 0.5
End Function
Public Function GetPerimeter() As Single
    Dim I      As Long

    GetPerimeter = 0
    For I = 1 To NumS
        GetPerimeter = GetPerimeter + MySprings(I).Length
    Next
End Function
Private Function GetDIR() As tVector


End Function

Public Sub SetUpSprings(Mode As tShapeMode, Radial As Boolean)

    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    Dim I      As Long
    Dim J      As Long

    'Dim CurLength As Single

    Dim NP     As Long
    Dim NS     As Long


    NumS = NumP
    ReDim MySprings(1 To NumS)

    For I = 1 To NumP
        J = I + 1
        If J > NumP Then J = 1
        AddSpring I, I, J
    Next


    AREA = GetAREA
    Perimeter = GetPerimeter
    BodyCenter = GetBodyCenter(AREA)
    SpringDir = GetDIR


    For I = 1 To NumP
        Mypoints(I).CentDir.X = (Mypoints(I).POS.X - BodyCenter.X)    ' / Mypoints(I).CentDIST
        Mypoints(I).CentDir.Y = (Mypoints(I).POS.Y - BodyCenter.Y)    ' / Mypoints(I).CentDIST
    Next


    Dim Pre    As Long
    Dim Nex    As Long
    Dim vP     As tVector
    Dim vN     As tVector
    Dim NewP   As tVector



    '    '----------------------------------------------------------
    '    'ShapedString
    '    NS = NumS
    '    NP = NumP
    '    GoTo skip
    '    For I = 1 To NS
    '        Pre = IIf(I > 1, I - 1, NS)
    '        Nex = IIf(I < NS, I + 1, 1)
    '        vN.X = Mypoints(Nex).POS.X - Mypoints(I).POS.X
    '        vN.y = Mypoints(Nex).POS.y - Mypoints(I).POS.y
    '        vP.X = Mypoints(I).POS.X - Mypoints(Pre).POS.X
    '        vP.y = Mypoints(I).POS.y - Mypoints(Pre).POS.y
    '        VectorPerp vN
    '        VectorPerp vP
    '        NewP.X = vN.X + vP.X
    '        NewP.y = vN.y + vP.y
    '        NewP = VectorNormalize(NewP)
    '        NewP.X = Mypoints(I).POS.X + NewP.X * Spessore
    '        NewP.y = Mypoints(I).POS.y + NewP.y * Spessore
    '        AddPoint NewP.X, NewP.y
    '    Next
    '
    'skip:
    '    Dim I2 As Long
    '    Dim J2 As Long
    '
    '    GoTo forbox
    '    '----------------------------------------------------------
    '    'For BALL
    '    For I = 1 To NP
    '        I2 = IIf(I < NP, I + 1, 1)
    '        J = I + NP
    '        J2 = IIf(J < NP * 2, J + 1, NP + 1)
    '        AddShapeSpring J, J2
    '        AddShapeSpring J, I2
    '        AddShapeSpring I, J2
    '        AddShapeSpring I, J
    '
    '
    '        'pre = IIf(I > 1, I - 1, NS)
    '        'nex = IIf(I < NS, I + 1, 1)
    '        'AddShapeSpring pre, nex
    '    Next
    '    GoTo FINE
    '
    'forbox:
    '
    '   For I = 2 To NP \ 2 - 1
    '        AddShapeSpring I, NP - I + 1
    '    Next
    '
    '

    If Mode = ByInternalSprings And Radial Then
        '    Stop

        AddPoint BodyCenter.X, BodyCenter.Y

        For I = 1 To NumS
            AddShapeSpring I, NumP
        Next


    End If

End Sub

Public Sub Create_BALL(pX, pY, BALLRADIUS, points, Smode As tShapeMode)
    Dim I      As Long

    Me_Kill

    ShapeMode = Smode


    For I = 1 To points
        AddPoint BALLRADIUS * Cos((2 * PI) * I / points) + pX, BALLRADIUS * Sin((2 * PI) * I / points) + pY
    Next

    SetUpSprings Smode, True

    AddALLMuscles

    'InflatedGas = Sqr(AREA)
    InflatedGas = 1.5 * AREA / Perimeter


End Sub
Public Sub Create_BOX(pX, pY, WW, HH, Smode As tShapeMode)

    Me_Kill

    AddPoint pX, pY
    AddPoint pX + WW, pY
    AddPoint pX + WW, pY + HH
    AddPoint pX, pY + HH
    ShapeMode = Smode
    SetUpSprings Smode, False
    If Smode = ByInternalSprings Then
        AddShapeSpring 1, 3
        AddShapeSpring 2, 4
    End If




    AddALLMuscles
    InflatedGas = 1.5 * AREA / Perimeter
End Sub

Public Sub Create_Rectangle(pX, pY, LinkLen, divX, DivY, Smode As tShapeMode)

    Dim X      As Single
    Dim Y      As Single

    Me_Kill

    ShapeMode = Smode



    Debug.Print "**************************"

    For X = 0 To divX
        AddPoint pX + (X) * LinkLen, pY
        Debug.Print Mypoints(NumP).POS.X - pX, Mypoints(NumP).POS.Y - pY
    Next
    For Y = 1 To DivY
        AddPoint pX + (divX) * LinkLen, pY + (Y) * LinkLen
        Debug.Print Mypoints(NumP).POS.X - pX, Mypoints(NumP).POS.Y - pY
    Next
    For X = divX - 1 To 0 Step -1
        AddPoint pX + (X) * LinkLen, pY + (DivY) * LinkLen
        Debug.Print Mypoints(NumP).POS.X - pX, Mypoints(NumP).POS.Y - pY
    Next
    For Y = DivY - 1 To 1 Step -1
        AddPoint pX, pY + (Y) * LinkLen
        Debug.Print Mypoints(NumP).POS.X - pX, Mypoints(NumP).POS.Y - pY
    Next

    'Stop


    SetUpSprings Smode, False

    AddALLMuscles

    InflatedGas = 1.5 * AREA / Perimeter

End Sub
Public Sub AddALLMuscles()
    Dim I      As Long
    For I = 1 To NumS - 1
        ADDMuscle I, I + 1, 100    '80
    Next
    ADDMuscle NumS, 1, 100    '80
End Sub

Public Sub Do_Gravity()
'___________________________________________________________________________
'************************
'* gravity */
'************************
    Dim I      As Long

    For I = 1 To NumP
        If Mypoints(I).Mass <> PositiveInfinity Then
            Mypoints(I).F.X = 0
            Mypoints(I).F.Y = Mypoints(I).Mass * GY
        End If
    Next

End Sub
Public Sub DO_FORCES()
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single

    Dim VX12   As Single
    Dim VY12   As Single

    Dim F      As Single
    Dim fX     As Single
    Dim fY     As Single

    Dim I      As Long


    Dim PressureV As Single





    '___________________________________________________________________________
    '************************
    'Spring Linear Force
    '************************

    For I = 1 To NumS

        '// get positions of spring start & end points
        X1 = Mypoints(MySprings(I).P1).POS.X
        Y1 = Mypoints(MySprings(I).P1).POS.Y
        X2 = Mypoints(MySprings(I).P2).POS.X
        Y2 = Mypoints(MySprings(I).P2).POS.Y

        '  // calculate sqr(distance)
        MySprings(I).CurLength = Distance((X1 - X2), (Y1 - Y2))

        If MySprings(I).CurLength <> 0 Then

            '  // get velocities of start & end points
            VX12 = Mypoints(MySprings(I).P1).VEL.X - Mypoints(MySprings(I).P2).VEL.X
            VY12 = Mypoints(MySprings(I).P1).VEL.Y - Mypoints(MySprings(I).P2).VEL.Y

            '  // calculate force value
            '   f = (CurLength -  mySprings[i].length) * KStiffness +
            '       (vx12 * (x1 - x2) +
            '        vy12 * (y1 - y2)) * KDamping / CurLength;
            '
            F = (MySprings(I).CurLength - MySprings(I).Length) * KStiffness + _
                (VX12 * (X1 - X2) + _
                 VY12 * (Y1 - Y2)) * KDamping / MySprings(I).CurLength
            '    If F <> 0 Then Stop

            '  // force vector
            fX = ((X1 - X2) / MySprings(I).CurLength) * F
            fY = ((Y1 - Y2) / MySprings(I).CurLength) * F

            '  // accumulate force for starting point
            Mypoints(MySprings(I).P1).F.X = Mypoints(MySprings(I).P1).F.X - fX
            Mypoints(MySprings(I).P1).F.Y = Mypoints(MySprings(I).P1).F.Y - fY


            '  // accumulate force for end point
            Mypoints(MySprings(I).P2).F.X = Mypoints(MySprings(I).P2).F.X + fX
            Mypoints(MySprings(I).P2).F.Y = Mypoints(MySprings(I).P2).F.Y + fY

        End If

        '  // Calculate normal vectors to springs
        If MySprings(I).CurLength <> 0 Then
            MySprings(I).Normal.X = (Y1 - Y2) / MySprings(I).CurLength
            MySprings(I).Normal.Y = -(X1 - X2) / MySprings(I).CurLength
        End If

    Next

    'Stop

    '___________________________________________________________________________
    '*******************************
    If ShapeDO Then

        If ShapeMode = ByMuscle Then

            DoMuscles

        Else
            '        Stop

            '************************
            'Shape Spring Linear Force
            '************************
            Dim Pre As Long
            Dim Nex As Long
            Dim vN As tVector
            Dim vP As tVector
            Dim NewP As tVector

            'For I = 1 To NumS
            '    Pre = IIf(I > 1, I - 1, NumS)
            '    Nex = IIf(I < NumS, I + 1, 1)
            '    vN.x = Mypoints(Nex).POS.x - Mypoints(I).POS.x
            '    vN.y = Mypoints(Nex).POS.y - Mypoints(I).POS.y
            '    vP.x = Mypoints(I).POS.x - Mypoints(Pre).POS.x
            '    vP.y = Mypoints(I).POS.y - Mypoints(Pre).POS.y
            '
            '    VectorPerp vN
            '    VectorPerp vP
            '
            '    NewP.x = vN.x + vP.x
            '    NewP.y = vN.y + vP.y
            '
            '    NewP = VectorNormalize(NewP)
            ''  Stop
            '
            '    Mypoints(I + NumS).POS.x = Mypoints(I).POS.x + NewP.x * 20 * SgnAREA ' Spessore
            '    Mypoints(I + NumS).POS.y = Mypoints(I).POS.y + NewP.y * 20 * SgnAREA 'Spessore
            '    Mypoints(I + NumS).VEL.x = 0
            '    Mypoints(I + NumS).VEL.y = 0
            '    Mypoints(I + NumS).F.x = 0
            '    Mypoints(I + NumS).F.y = 0
            '
            'Next
            '
            For I = 1 To NumSS

                '// get positions of spring start & end points
                X1 = Mypoints(ShapeSprings(I).P1).POS.X
                Y1 = Mypoints(ShapeSprings(I).P1).POS.Y
                X2 = Mypoints(ShapeSprings(I).P2).POS.X
                Y2 = Mypoints(ShapeSprings(I).P2).POS.Y

                '  // calculate sqr(distance)
                ShapeSprings(I).CurLength = Distance((X1 - X2), (Y1 - Y2))

                If ShapeSprings(I).CurLength <> 0 Then

                    '  // get velocities of start & end points
                    VX12 = Mypoints(ShapeSprings(I).P1).VEL.X - Mypoints(ShapeSprings(I).P2).VEL.X
                    VY12 = Mypoints(ShapeSprings(I).P1).VEL.Y - Mypoints(ShapeSprings(I).P2).VEL.Y

                    '  // calculate force value
                    '   f = (CurLength -  ShapeSprings[i].length) * KStiffness +
                    '       (vx12 * (x1 - x2) +
                    '        vy12 * (y1 - y2)) * KDamping / CurLength;
                    '
                    F = (ShapeSprings(I).CurLength - ShapeSprings(I).Length) * KStiffness + _
                        (VX12 * (X1 - X2) + _
                         VY12 * (Y1 - Y2)) * KDamping / ShapeSprings(I).CurLength
                    '    If F <> 0 Then Stop

                    '  // force vector

                    fX = ((X1 - X2) / ShapeSprings(I).CurLength) * F * 0.5
                    fY = ((Y1 - Y2) / ShapeSprings(I).CurLength) * F * 0.5


                    '*****************************************************
                    '  // accumulate force for starting point
                    Mypoints(ShapeSprings(I).P1).F.X = Mypoints(ShapeSprings(I).P1).F.X - fX
                    Mypoints(ShapeSprings(I).P1).F.Y = Mypoints(ShapeSprings(I).P1).F.Y - fY


                    '  // accumulate force for end point
                    Mypoints(ShapeSprings(I).P2).F.X = Mypoints(ShapeSprings(I).P2).F.X + fX
                    Mypoints(ShapeSprings(I).P2).F.Y = Mypoints(ShapeSprings(I).P2).F.Y + fY

                End If

                '  // Calculate normal vectors to springs
                ShapeSprings(I).Normal.X = (Y1 - Y2) / ShapeSprings(I).CurLength
                ShapeSprings(I).Normal.Y = -(X1 - X2) / ShapeSprings(I).CurLength
            Next
        End If



        '*******************************
    End If




    '___________________________________________________________________________
    '**********************
    'CurrAREA of The Body
    '**********************
    CurrAREA = GetAREA
    SgnAREA = Sgn(CurrAREA)
    CurrBodyCenter = GetBodyCenter(CurrAREA)
    CurrAREA = Abs(CurrAREA)
    CurrSpringDir = GetDIR


    '___________________________________________________________________________
    '********************
    'Pressure Force Distribution
    '********************


    For I = 1 To NumS
        With MySprings(I)
            X1 = Mypoints(.P1).POS.X
            Y1 = Mypoints(.P1).POS.Y
            X2 = Mypoints(.P2).POS.X
            Y2 = Mypoints(.P2).POS.Y

            ' PressureV = CurLength * AREA * (1 / CurrAREA)
            'PressureV = KPressure * CurLength * (AREA + (AREA - CurrAREA))
            PressureV = MySprings(I).CurLength * (InflatedGas) * (AREA / CurrAREA)


            Mypoints(.P1).F.X = Mypoints(.P1).F.X + MySprings(I).Normal.X * PressureV
            Mypoints(.P1).F.Y = Mypoints(.P1).F.Y + MySprings(I).Normal.Y * PressureV
            Mypoints(.P2).F.X = Mypoints(.P2).F.X + MySprings(I).Normal.X * PressureV
            Mypoints(.P2).F.Y = Mypoints(.P2).F.Y + MySprings(I).Normal.Y * PressureV
        End With
    Next



    '___________________________________________________________________________
    '*******************************
    'Integrate Newton 's Equations
    '*******************************
    For I = 1 To NumP

        If Mypoints(I).Mass = PositiveInfinity Then
            Mypoints(I).VEL.X = 0
            Mypoints(I).VEL.Y = 0
            Mypoints(I).F.X = 0
            Mypoints(I).F.Y = 0

        End If

        Mypoints(I).VEL.X = Mypoints(I).VEL.X + (Mypoints(I).F.X / Mypoints(I).Mass) * DT
        Mypoints(I).VEL.Y = Mypoints(I).VEL.Y + (Mypoints(I).F.Y / Mypoints(I).Mass) * DT

        Mypoints(I).POS.X = Mypoints(I).POS.X + Mypoints(I).VEL.X * DT
        Mypoints(I).POS.Y = Mypoints(I).POS.Y + Mypoints(I).VEL.Y * DT

        Mypoints(I).VEL.X = Mypoints(I).VEL.X * AirResistence
        Mypoints(I).VEL.Y = Mypoints(I).VEL.Y * AirResistence

        Mypoints(I).CpyPOS.X = Mypoints(I).POS.X
        Mypoints(I).CpyPOS.Y = Mypoints(I).POS.Y

        Mypoints(I).CpyVEL.X = Mypoints(I).VEL.X
        Mypoints(I).CpyVEL.Y = Mypoints(I).VEL.Y
    Next

    UpdateAABB


End Sub

Sub Do_ScreenBoundaries()
    Dim I      As Long

    ''  keep soft body within screen boundaries
    For I = 1 To NumP
        With Mypoints(I)
            ''  right
            If .POS.X > PicW - 1 Then
                .VEL.X = -.VEL.X
                .POS.X = PicW - 1
                .VEL.X = .VEL.X * Wall_Friction
                .VEL.Y = .VEL.Y * Wall_Friction
            End If
            ''  left
            If .POS.X < 1 Then
                .VEL.X = -.VEL.X
                .POS.X = 1
                .VEL.X = .VEL.X * Wall_Friction
                .VEL.Y = .VEL.Y * Wall_Friction
            End If
            ''  bottom
            If .POS.Y > PicH - 2 Then

                '      Stop

                .VEL.Y = -.VEL.Y
                .POS.Y = PicH - 2
                .VEL.X = .VEL.X * Wall_Friction
                .VEL.Y = .VEL.Y * Wall_Friction
            End If
            ''  top
            If .POS.Y < 1 Then
                .VEL.Y = -.VEL.Y
                .POS.Y = 1
                .VEL.X = .VEL.X * Wall_Friction
                .VEL.Y = .VEL.Y * Wall_Friction
            End If
        End With
    Next
End Sub

Public Sub Me_DRAW(PicHdc As Long)
    Dim I      As Long

    Dim X1     As Long
    Dim Y1     As Long
    Dim X2     As Long
    Dim Y2     As Long
    Dim dCOLOR As Long

    Dim PRESS  As Single

    Dim PenSelect As Long
    Dim BrushSelect As Long

    Dim PP()   As POINTAPI

    ReDim PP(0 To NumS)

    PRESS = InflatedGas * (AREA / CurrAREA)


    For I = 1 To NumS


        With MySprings(I)
            X1 = Mypoints(.P1).POS.X
            Y1 = Mypoints(.P1).POS.Y


            PP(I).X = X1
            PP(I).Y = Y1


            If Mypoints(.P1).Mass = PositiveInfinity Then MyCircle PicHdc, X1, Y1, 3, 3, vbWhite

        End With
    Next

    If PRESS > 100 Then
        PRESS = 255
    Else
        PRESS = PRESS * 20
        If PRESS > 255 Then PRESS = 255
    End If

    'dCOLOR = RGB(PRESS, 255 - PRESS * 0.5, 0)
    dCOLOR = COLOR


    PenSelect = SelectObject(PicHdc, CreatePen(0, 1, vbYellow))    'RGB(200, 200, 200)))
    BrushSelect = SelectObject(PicHdc, CreateSolidBrush(dCOLOR))

    PP(0) = PP(NumS)
    Polygon PicHdc, PP(0), NumS
    DeleteObject PenSelect
    DeleteObject BrushSelect



    If ShapeMode = ByInternalSprings And ShapeDO Then

        For I = 1 To NumSS

            With ShapeSprings(I)
                X1 = Mypoints(.P1).POS.X
                Y1 = Mypoints(.P1).POS.Y
                X2 = Mypoints(.P2).POS.X
                Y2 = Mypoints(.P2).POS.Y

                FastLine PicHdc, X1, Y1, X2, Y2, 1, RGB(200, 200, 200)

            End With
        Next
    End If

End Sub


Public Sub DO_Shape()
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    Dim I      As Long
    X1 = CurrBodyCenter.X
    Y1 = CurrBodyCenter.Y

    For I = 1 To NumP
        'If I \ 2 = I / 2 Then
        X2 = X1 - (CurrSpringDir.X * Mypoints(I).CentDir.X - CurrSpringDir.Y * Mypoints(I).CentDir.Y)
        Y2 = Y1 - (CurrSpringDir.Y * Mypoints(I).CentDir.X + CurrSpringDir.X * Mypoints(I).CentDir.Y)

        'Mypoints(I).F.x = Mypoints(I).F.x + (X2 - Mypoints(I).POS.x) * InflatedGas * 0.5
        'Mypoints(I).F.y = Mypoints(I).F.y + (Y2 - Mypoints(I).POS.y) * InflatedGas * 0.5
        'Mypoints(I).F.x = Mypoints(I).F.x + Sgn(X2 - Mypoints(I).POS.x) * 200
        'Mypoints(I).F.y = Mypoints(I).F.y + Sgn(Y2 - Mypoints(I).POS.y) * 200

        Mypoints(I).F.X = Mypoints(I).F.X + (X2 - Mypoints(I).POS.X) * 20
        Mypoints(I).F.Y = Mypoints(I).F.Y + (Y2 - Mypoints(I).POS.Y) * 20

        'End If

    Next

End Sub

Sub DoMuscles()
    Dim X1     As Single
    Dim Y1     As Single
    Dim X2     As Single
    Dim Y2     As Single
    Dim AM     As Single
    Dim Dx1    As Single
    Dim Dy1    As Single
    Dim Dx2    As Single
    Dim Dy2    As Single

    Dim RetP   As tVector

    Dim Perc1  As Single
    Dim Perc2  As Single

    Dim M      As Long

    Dim A1     As Single
    Dim A2     As Single
    Dim D      As Single

    'Const Kfixed = 0.1
    Const kMuscleSpeedLimit = 80    '80 ' 50 ' 10 '10

    For M = 1 To NMuscles
        With Muscle(M)
            'MyPoints(.P0).NewX = MyPoints(.P0).x + MyPoints(.P0).f.x
            'MyPoints(.P0).NewY = MyPoints(.P0).y + MyPoints(.P0).f.Y
            'MyPoints(.P1).NewX = MyPoints(.P1).x + MyPoints(.P1).f.x
            'MyPoints(.P1).NewY = MyPoints(.P1).y + MyPoints(.P1).f.Y
            'MyPoints(.P2).NewX = MyPoints(.P2).x + MyPoints(.P2).f.x
            'MyPoints(.P2).NewY = MyPoints(.P2).y + MyPoints(.P2).f.Y
        End With
    Next

    For M = 1 To NMuscles

        With Muscle(M)

            A1 = Atan2(Mypoints(.P1).POS.X - Mypoints(.P0).POS.X, Mypoints(.P1).POS.Y - Mypoints(.P0).POS.Y)
            A2 = Atan2(Mypoints(.P2).POS.X - Mypoints(.P0).POS.X, Mypoints(.P2).POS.Y - Mypoints(.P0).POS.Y)

            '            AM = (A1 - A2) - .MainA
            AM = (A1 - A2) + .MainA


            'If .IsDynamic Then                                  'pi/180
            '    AM = AM + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp * 0.017453292
            'End If


            '--------------------------------------------------
            'This Should Make Longer Link Move Slower
            'and shorter move faster
            X1 = Mypoints(.P0).POS.X + Cos(A1 - AM) * MySprings(.L1).Length
            Y1 = Mypoints(.P0).POS.Y + Sin(A1 - AM) * MySprings(.L1).Length
            X2 = Mypoints(.P0).POS.X + Cos(A2 + AM) * MySprings(.L2).Length
            Y2 = Mypoints(.P0).POS.Y + Sin(A2 + AM) * MySprings(.L2).Length
            '
            '            If MySprings(.L1).CurLength > MySprings(.L2).CurLength Then
            '                Perc1 = MySprings(.L2).CurLength / MySprings(.L1).CurLength '(MySprings(.L1).CurrLen + MySprings(.L2).CurrLen)
            '                Perc2 = 1 - Perc1
            '            Else '
            '                Perc2 = MySprings(.L1).CurLength / MySprings(.L2).CurLength '(MySprings(.L1).CurrLen + MySprings(.L2).CurrLen)
            '                Perc1 = 1 - Perc2
            '            End If
            '--------------------------------------------------
            Perc1 = 0.5
            Perc2 = 0.5
            '            Stop

            Dx1 = (Mypoints(.P1).POS.X - X1) * Perc1
            Dy1 = (Mypoints(.P1).POS.Y - Y1) * Perc1
            Dx2 = (Mypoints(.P2).POS.X - X2) * Perc2
            Dy2 = (Mypoints(.P2).POS.Y - Y2) * Perc2

            GoTo SkipLimit
            'Apply OK Speed Limit
            D = Sqr(Dx1 * Dx1 + Dy1 * Dy1)
            If D > kMuscleSpeedLimit Then
                RetP = VectorNormalize(Vector(Dx1, Dy1))
                Dx1 = RetP.X * kMuscleSpeedLimit
                Dy1 = RetP.Y * kMuscleSpeedLimit
            End If
            D = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
            If D > kMuscleSpeedLimit Then
                RetP = VectorNormalize(Vector(Dx2, Dy2))
                Dx2 = RetP.X * kMuscleSpeedLimit
                Dy2 = RetP.Y * kMuscleSpeedLimit
            End If

SkipLimit:


            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            '            If .FixedANG Then
            '                MyPoints(.P1).f.x = MyPoints(.P1).f.x - Kfixed * Dx1
            '                MyPoints(.P1).f.Y = MyPoints(.P1).f.Y - Kfixed * Dy1
            '                MyPoints(.P2).f.x = MyPoints(.P2).f.x - Kfixed * Dx2
            '                MyPoints(.P2).f.Y = MyPoints(.P2).f.Y - Kfixed * Dy2
            '                MyPoints(.P0).f.x = MyPoints(.P0).f.x + Kfixed * (Dx1 + Dx2)
            '                MyPoints(.P0).f.Y = MyPoints(.P0).f.Y + Kfixed * (Dy1 + Dy2)
            '            Else

            Mypoints(.P1).F.X = Mypoints(.P1).F.X - .F * Dx1
            Mypoints(.P1).F.Y = Mypoints(.P1).F.Y - .F * Dy1

            Mypoints(.P2).F.X = Mypoints(.P2).F.X - .F * Dx2
            Mypoints(.P2).F.Y = Mypoints(.P2).F.Y - .F * Dy2

            Mypoints(.P0).F.X = Mypoints(.P0).F.X + .F * (Dx1 + Dx2)
            Mypoints(.P0).F.Y = Mypoints(.P0).F.Y + .F * (Dy1 + Dy2)

            '    MyPoints(.P0).f.x = MyPoints(.P0).f.x + .f * Dx1
            '    MyPoints(.P0).f.Y = MyPoints(.P0).f.Y + .f * Dy1

            '            End If

        End With

    Next

End Sub
Private Sub UpdateAABB()
    Dim I      As Long


    AABB.Max.X = -9999999
    AABB.Max.Y = -9999999
    AABB.Min.X = 9999999
    AABB.Min.Y = 9999999

    For I = 1 To NumP

        If Mypoints(I).POS.X > AABB.Max.X Then AABB.Max.X = Mypoints(I).POS.X
        If Mypoints(I).POS.Y > AABB.Max.Y Then AABB.Max.Y = Mypoints(I).POS.Y

        If Mypoints(I).POS.X < AABB.Min.X Then AABB.Min.X = Mypoints(I).POS.X
        If Mypoints(I).POS.Y < AABB.Min.Y Then AABB.Min.Y = Mypoints(I).POS.Y

    Next


End Sub


'collision detection.  detect if a global point is inside this body.
'<returns>true = point is inside body, false = it is not.</returns>
Public Function Contains(X, Y) As Boolean
    Dim Pt     As tVector
    Dim I      As Long

    Pt.X = X
    Pt.Y = Y

    Dim Inside As Boolean
    Dim EndPt  As tVector
    Dim EdgeSt As tVector
    Dim EdgeEnd As tVector
    Dim Slope  As Single
    Dim HitX   As Single
    Dim C      As Long
    C = NumP

    '// basic idea: draw a line from the point to a point known to be outside the body.  count the number of
    '// lines in the polygon it intersects.  if that number is odd, we are inside.  if it's even, we are outside.
    '// in this implementation we will always use a line that moves off in the positive X direction from the point
    '// to simplify things.

    EndPt.X = AABB.Max.X + 0.1
    EndPt.Y = Pt.Y

    'line we are testing against goes from pt -> endPt.
    Inside = False
    EdgeSt = Mypoints(1).POS

    For I = 1 To C

        'the current edge is defined as the line from edgeSt -> edgeEnd.
        If (I < (C)) Then
            EdgeEnd = Mypoints(I + 1).POS
        Else
            EdgeEnd = Mypoints(1).POS
        End If
        'perform check now...
        If (((EdgeSt.Y <= Pt.Y) And (EdgeEnd.Y > Pt.Y)) Or ((EdgeSt.Y > Pt.Y) And (EdgeEnd.Y <= Pt.Y))) Then
            'this line crosses the test line at some point... does it do so within our test range?
            Slope = (EdgeEnd.X - EdgeSt.X) / (EdgeEnd.Y - EdgeSt.Y)
            HitX = EdgeSt.X + ((Pt.Y - EdgeSt.Y) * Slope)

            If ((HitX >= Pt.X) And (HitX <= EndPt.X)) Then Inside = Not (Inside)
        End If


        EdgeSt = EdgeEnd

    Next


    Contains = Inside

End Function

'collision detection - given a global point, find the point on this body that is closest to the global point,
'and if it is an edge, information about the edge it resides on.
'"pt">global point</param>
'"hitPt">returned point on the body in global space</param>
'"normal">returned normal on the body in global space</param>
'"pointA">returned ptA on the edge</param>
'"pointB">returned ptB on the edge</param>
'"edgeD">scalar distance between ptA and ptB [0,1]</param>
'distance</returns>

Public Function GetClosestPoint(ByVal ptX, ByVal ptY, ByRef hitPtX, ByRef hitPtY, ByRef NormalX, ByRef NormalY, ByRef PointA, ByRef PointB, ByRef EdgeD) As Single
'public float getClosestPoint( Vector2 pt, out Vector2 hitPt, out Vector2 normal, out int pointA, out int pointB, out float edgeD )

    Dim HitPt  As tVector
    Dim Pt     As tVector
    Dim Normal As tVector
    Dim tempHit As tVector
    Dim tempNorm As tVector
    Dim tempEdgeD As Single
    Dim closestD As Single
    Dim DIST   As Single
    Dim I      As Long

    Pt.X = ptX
    Pt.Y = ptY
    HitPt.X = 0
    HitPt.Y = 0
    PointA = -1
    PointB = -1
    EdgeD = 0
    Normal.X = 0
    Normal.Y = 0

    closestD = 1000000    '1000


    For I = 1 To NumP


        'float dist = getClosestPointOnEdge(pt, i, out tempHit, out tempNorm, out tempEdgeD);
        DIST = GetClosestPointOnEdge(Pt.X, Pt.Y, I, tempHit.X, tempHit.Y, tempNorm.X, tempNorm.Y, tempEdgeD)
        If (DIST < closestD) Then

            closestD = DIST
            PointA = I
            If (I < NumP) Then
                PointB = I + 1
            Else
                PointB = 1
            End If

            EdgeD = tempEdgeD
            Normal = tempNorm
            HitPt = tempHit
        End If
    Next

    hitPtX = HitPt.X
    hitPtY = HitPt.Y
    NormalX = Normal.X
    NormalY = Normal.Y


    '// return.
    GetClosestPoint = closestD


End Function



'find the distance from a global point in space, to the closest point on a given edge of the body.
'"pt">global point</param>
'"edgeNum">edge to check against.  0 = edge from pt[0] to pt[1], etc.</param>
'"hitPt">returned point on edge in global space</param>
'"normal">returned normal on edge in global space</param>
'"edgeD">returned distance along edge from ptA to ptB [0,1]</param>
'<returns>distance</returns>

'        public float getClosestPointOnEdge(Vector2 pt, int edgeNum, out Vector2 hitPt, out Vector2 normal, out float edgeD)
Public Function GetClosestPointOnEdge(ByVal ptX As Single, ByVal ptY As Single, ByVal EdgeNum, ByRef hitPtX, ByRef hitPtY, ByRef NormalX, ByRef NoramlY, ByRef EdgeD) As Single
    Dim Pt     As tVector
    Dim HitPt  As tVector
    Dim Normal As tVector
    Dim DIST   As Single
    Dim ptA    As tVector
    Dim ptB    As tVector
    Dim toP    As tVector
    Dim E      As tVector
    Dim edgeLength As Single
    Dim N      As tVector
    Dim TmpE   As tVector
    Dim X      As Single
    Dim toP3   As tVector3
    Dim E3     As tVector3

    Pt.X = ptX
    Pt.Y = ptY


    HitPt.X = 0
    HitPt.Y = 0

    Normal.X = 0
    Normal.Y = 0

    EdgeD = 0
    DIST = 0

    ptA = Mypoints(EdgeNum).POS
    'Vector2 ptB = new Vector2();

    If (EdgeNum < (NumP)) Then
        ptB = Mypoints(EdgeNum + 1).POS
    Else
        ptB = Mypoints(1).POS
    End If

    toP.X = Pt.X - ptA.X
    toP.Y = Pt.Y - ptA.Y


    E.X = ptB.X - ptA.X
    E.Y = ptB.Y - ptA.Y

    'get the length of the edge, and use that to normalize the vector.
    edgeLength = Sqr((E.X * E.X) + (E.Y * E.Y))
    If (edgeLength > 0.00001) Then
        E.X = E.X / edgeLength
        E.Y = E.Y / edgeLength
    End If

    'normal
    'VectorTools.getPerpendicular(ref E, ref n);
    TmpE = E
    VectorPerp TmpE
    N = TmpE

    'calculate the distance!
    'Vector2.Dot(ref toP, ref E, out x);
    X = VectorDot(toP, E)

    If (X <= 0) Then

        'x is outside the line segment, distance is from pt to ptA.
        '//dist = (pt - ptA).Length();
        DIST = VectorDist(Pt, ptA)
        'Vector2.Distance(ref pt, ref ptA, out dist);

        HitPt = ptA
        EdgeD = 0
        Normal = N


    ElseIf (X >= edgeLength) Then

        'x is outside of the line segment, distance is from pt to ptB.
        'dist = (pt - ptB).Length();
        'Vector2.Distance(ref pt, ref ptB, out dist);
        DIST = VectorDist(Pt, ptB)
        HitPt = ptB
        EdgeD = 1
        Normal = N



    Else

        ' point lies somewhere on the line segment.

        toP3.X = toP.X
        toP3.Y = toP.Y
        toP3.Z = 0    '1????????

        E3.X = E.X
        E3.Y = E.Y
        E3.Z = 0    '1?????????????

        '//dist = Math.Abs(Vector3.Cross(toP3, E3).Z);
        '
        DIST = VectorCross3(toP3, E3).Z

        'Vector3.Cross(ref toP3, ref E3, out E3);
        E3 = VectorCross3(toP3, E3)

        DIST = Abs(E3.Z)

        HitPt.X = ptA.X + (E.X * X)
        HitPt.Y = ptA.Y + (E.Y * X)
        EdgeD = X / edgeLength
        Normal = N
    End If

    '            return dist;
    GetClosestPointOnEdge = DIST

    hitPtX = HitPt.X
    hitPtY = HitPt.Y
    NormalX = Normal.X
    NoramlY = Normal.Y


End Function



'Helper function to add a global force acting on this body as a whole.
'"pt">location of force, in global space</param>
'"force">direction and intensity of force, in global space</param>

Public Sub AddGlobalForce(ptX, ptY, ForceX, ForceY)
    Dim Pt     As tVector
    Dim Force  As tVector
    Dim R      As tVector
    Dim Force3 As tVector3
    Dim R3     As tVector3
    Dim toPT   As tVector
    Dim Torque As tVector
    Dim TorqueF As Single

    Dim I      As Long

    Pt.X = ptX
    Pt.Y = ptY
    Force.X = ForceX
    Force.Y = ForceY


    'Vector2 R = (mDerivedPos - pt);
    R.X = BodyCenter.X - Pt.X
    R.Y = BodyCenter.Y - Pt.Y


    R3.X = R.X
    R3.Y = R.Y
    R3.Z = 0    '1'?????????????

    Force3.X = Force.X
    Force3.Y = Force.Y
    Force3.Z = 0    '1'?????????????

    'float torqueF = Vector3.Cross(JelloPhysics.VectorTools.vec3FromVec2(R), JelloPhysics.VectorTools.vec3FromVec2(force)).Z;
    TorqueF = VectorCross3(R3, Force3).Z

    For I = 1 To NumP

        'Vector2 toPt = (mPointMasses[i].Position - mDerivedPos);
        'Vector2 torque = JelloPhysics.VectorTools.rotateVector(toPt, -(float)(Math.PI) / 2f);
        toPT.X = Mypoints(I).POS.X - BodyCenter.X
        toPT.Y = Mypoints(I).POS.Y - BodyCenter.Y
        Torque = RotatedVector(toPT, -PI / 2)

        'mPointMasses[i].Force += torque * torqueF;
        'mPointMasses[i].Force += force;

        Mypoints(I).F.X = Mypoints(I).F.X = Torque.X * TorqueF
        Mypoints(I).F.Y = Mypoints(I).F.Y = Torque.Y * TorqueF

        Mypoints(I).F.X = Mypoints(I).F.X = Force.X
        Mypoints(I).F.Y = Mypoints(I).F.Y = Force.Y

    Next
End Sub


Public Function getClosestPointOnEdgeSquared(ptX, ptY, EdgeNum, ByRef hitPtX, ByRef hitPtY, ByRef NormalX, ByRef NormalY, ByRef EdgeD) As Single
'public float getClosestPointOnEdgeSquared(Vector2 pt, int edgeNum, out Vector2 hitPt, out Vector2 normal, out float edgeD)
    Dim HitPt  As tVector
    Dim Pt     As tVector
    Dim Normal As tVector
    Dim DIST   As Single

    Dim ptA    As tVector
    Dim ptB    As tVector

    Dim toP    As tVector
    Dim E      As tVector
    Dim toP3   As tVector3
    Dim E3     As tVector3

    Dim edgeLength As Single

    Dim N      As tVector

    Dim X      As Single

    Pt.X = ptX
    Pt.Y = ptY

    'hitPt = new Vector2();
    'hitPt.X = 0f;
    'hitPt.Y = 0f;

    'normal = new Vector2();
    'normal.X = 0f;
    'normal.Y = 0f;

    'edgeD = 0f;
    'float dist = 0f;

    HitPt.X = 0
    HitPt.Y = 0
    Normal.X = 0
    Normal.Y = 0
    EdgeD = 0
    DIST = 0


    'Vector2 ptA = mPointMasses[edgeNum].Position;
    'Vector2 ptB = new Vector2();

    ptA = Mypoints(EdgeNum).POS


    '            if (edgeNum < (mPointMasses.Count - 1))
    '                ptB = mPointMasses[edgeNum + 1].Position;
    '            Else
    '                ptB = mPointMasses[0].Position;

    If (EdgeNum < NumP) Then
        ptB = Mypoints(EdgeNum + 1).POS
    Else
        ptB = Mypoints(1).POS
    End If

    'Vector2 toP = new Vector2();
    'toP.X = pt.X - ptA.X;
    'toP.Y = pt.Y - ptA.Y;
    toP.X = Pt.X - ptA.X
    toP.Y = Pt.Y - ptA.Y

    'Vector2 E = new Vector2();
    'E.X = ptB.X - ptA.X;
    'E.Y = ptB.Y - ptA.Y;
    E.X = ptB.X - ptA.X
    E.Y = ptB.Y - ptA.Y


    '// get the length of the edge, and use that to normalize the vector.
    'float edgeLength = (float)Math.Sqrt((E.X * E.X) + (E.Y * E.Y));
    edgeLength = Sqr((E.X * E.X) + (E.Y * E.Y))


    'if (edgeLength > 0.00001f)
    '{
    '    E.X /= edgeLength;
    '    E.Y /= edgeLength;
    '}

    If (edgeLength > 0.00001) Then
        E.X = E.X / edgeLength
        E.Y = E.Y / edgeLength
    End If

    '// normal
    'Vector2 n = new Vector2();
    'VectorTools.getPerpendicular(ref E, ref n);
    N = E
    VectorPerp N

    '// calculate the distance!
    'float x;
    'Vector2.Dot(ref toP, ref E, out x);
    X = VectorDot(toP, E)


    'if (x <= 0.0f)
    If X <= 0 Then


        '{
        '    // x is outside the line segment, distance is from pt to ptA.
        '    //dist = (pt - ptA).Length();
        '    Vector2.DistanceSquared(ref pt, ref ptA, out dist);
        '    hitPt = ptA;
        '    edgeD = 0f;
        '    normal = n;
        '}
        DIST = VectorSquaredDist(Pt, ptA)


        HitPt = ptA
        EdgeD = 0
        Normal = N

        'else if (x >= edgeLength)
        '{
        '    // x is outside of the line segment, distance is from pt to ptB.
        '    //dist = (pt - ptB).Length();
        '    Vector2.DistanceSquared(ref pt, ref ptB, out dist);
        '    hitPt = ptB;
        '    edgeD = 1f;
        '    normal = n;
        '}
    ElseIf X >= edgeLength Then


        'DIST = VectorDist(Pt, ptB)
        DIST = VectorSquaredDist(Pt, ptB)

        HitPt = ptB
        EdgeD = 1
        Normal = N



        'Else
    Else
        '{
        '    // point lies somewhere on the line segment.
        '    Vector3 toP3 = new Vector3();
        '    toP3.X = toP.X;
        '    toP3.Y = toP.Y;

        '    Vector3 E3 = new Vector3();
        '    E3.X = E.X;
        '    E3.Y = E.Y;

        '    //dist = Math.Abs(Vector3.Cross(toP3, E3).Z);
        '    Vector3.Cross(ref toP3, ref E3, out E3);
        '    dist = Math.Abs(E3.Z * E3.Z);
        '    hitPt.X = ptA.X + (E.X * x);
        '    hitPt.Y = ptA.Y + (E.Y * x);
        '    edgeD = x / edgeLength;
        '    normal = n;
        '}


        toP3.X = toP.X
        toP3.Y = toP.Y
        toP3.Z = 0    '1'??????
        E3.X = E.X
        E3.Y = E.Y
        E3.Z = 0    '1'??????
        E3 = VectorCross3(toP3, E3)
        DIST = Abs(E3.Z * E3.Z)
        HitPt.X = ptA.X + (E.X * X)
        HitPt.Y = ptA.Y + (E.Y * X)
        EdgeD = X / edgeLength
        Normal = N
    End If

    'return dist;
    getClosestPointOnEdgeSquared = DIST

    hitPtX = HitPt.X
    hitPtY = HitPt.Y
    NormalX = Normal.X
    NormalY = Normal.Y

End Function

Private Function zR(S) As String
    zR = Replace(S, ",", ".")
End Function

Public Sub Me_Kill()
    Dim I      As Long
    For I = 1 To NumP
        Mypoints(I).VEL.X = 0
        Mypoints(I).VEL.Y = 0
        Mypoints(I).F.X = 0
        Mypoints(I).F.Y = 0
    Next

    NumP = 0
    NumS = 0
    NumSS = 0
    NMuscles = 0

End Sub



Public Function Me_SAVE(File As String)

    DO_FORCES

    Open App.Path & "\" & File & ".SoftBody" For Output As 1
    Dim X      As Single
    Dim Y      As Single
    Dim I      As Long

    X = CurrBodyCenter.X
    Y = CurrBodyCenter.Y

    Print #1, "------------------NumP-NumS-NUMSS-Kstif-Kdamp-InGas-Area-CurArea-ShapeMode-ShapeDO"

    Print #1, NumP
    Print #1, NumS
    Print #1, NumSS

    Print #1, zR(KStiffness)
    Print #1, zR(KDamping)
    Print #1, zR(InflatedGas)
    Print #1, zR(AREA)
    Print #1, zR(CurrAREA)
    Print #1, zR(Perimeter)
    Print #1, Material
    Print #1, COLOR

    Print #1, ShapeMode
    Print #1, CInt(ShapeDO)


    Print #1, "------------------POINTS  X  Y  Mass"
    For I = 1 To NumP
        Print #1, zR(Mypoints(I).POS.X - X)
        Print #1, zR(Mypoints(I).POS.Y - Y)
        Print #1, zR(Mypoints(I).Mass)
    Next
    Print #1, "------------------SPRINGS  P1 P2   Len   CurrLen"
    For I = 1 To NumS
        Print #1, MySprings(I).P1
        Print #1, MySprings(I).P2
        Print #1, zR(MySprings(I).Length)
        Print #1, zR(MySprings(I).CurLength)
    Next
    Print #1, "------------------SHAPE SPRINGS  P1 P2   Len   CurrLen"
    For I = 1 To NumSS
        Print #1, ShapeSprings(I).P1
        Print #1, ShapeSprings(I).P2
        Print #1, zR(ShapeSprings(I).Length)
        Print #1, zR(ShapeSprings(I).CurLength)
    Next
    Print #1, "------------------Muscles  P0 P1 P2  Stregth   (MainA)"
    For I = 1 To NMuscles
        Print #1, Muscle(I).P0
        Print #1, Muscle(I).P1
        Print #1, Muscle(I).P2
        Print #1, zR(Muscle(I).F)
        Print #1, zR(Muscle(I).MainA)
    Next

    Close 1

End Function

Public Function Me_LOAD(File As String, CposX, CposY)
    Dim S      As String
    Dim I      As Long
    Dim tNumP  As Long
    Dim tNumS  As Long
    Dim tNumSS As Long
    Dim Inte   As Integer
    Dim S1     As Single
    Dim S2     As Single
    Dim S3     As Single
    Dim S4     As Single

    Me_Kill

    Open App.Path & "\" & File & ".SoftBody" For Input As 1


    Input #1, S
    Input #1, tNumP
    Input #1, tNumS
    Input #1, tNumSS

    Input #1, KStiffness
    Input #1, KDamping
    Input #1, InflatedGas
    Input #1, AREA
    Input #1, CurrAREA
    Input #1, Perimeter
    Input #1, Material
    Input #1, COLOR

    Input #1, ShapeMode
    Input #1, Inte
    ShapeDO = CBool(Inte)

    '------------------------
    Input #1, S
    For I = 1 To tNumP
        Input #1, S1
        Input #1, S2
        Input #1, S3
        AddPoint S1 + CposX, S2 + CposY, S3
    Next
    '------------------------
    NumS = NumP
    ReDim MySprings(1 To NumS)
    '------------------------
    Input #1, S
    For I = 1 To tNumS
        Input #1, S1
        Input #1, S2
        Input #1, S3
        Input #1, S4
        AddSpring I, S1, S2
        MySprings(I).Length = S3
        MySprings(I).CurLength = S4
    Next
    '------------------------
    '------------------------
    Input #1, S
    For I = 1 To tNumSS
        Input #1, S1
        Input #1, S2
        Input #1, S3
        Input #1, S4
        AddShapeSpring S1, S2
        ShapeSprings(I).Length = S3
        ShapeSprings(I).CurLength = S4
    Next
    '------------------------


    Close 1
End Function
