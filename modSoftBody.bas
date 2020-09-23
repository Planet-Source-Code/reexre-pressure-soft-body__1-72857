Attribute VB_Name = "modsOFTbODY"
'---------------------------------------------------------------------------------------
' Module    :
' Author    : reexre@gmail.com Roberto Mior
' Date      : 22/01/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact

Option Explicit

Public Enum tShapeMode
    ByMuscle
    ByInternalSprings
End Enum

Public Type tVector
    X          As Single
    Y          As Single
End Type

Public Type tVector3
    X          As Single
    Y          As Single
    Z          As Single
End Type


Public Type tPoint

    VEL        As tVector
    POS        As tVector
    F          As tVector

    CentDir    As tVector

    CpyPOS     As tVector
    CpyVEL     As tVector

    Mass       As Single

End Type


Public Type tSpring

    P1         As Long
    P2         As Long
    Length     As Single
    Normal     As tVector
    CurLength  As Single

End Type


Public Type tMuscle

    L1         As Integer    '     Link1
    L2         As Integer    '     Link2
    MainA      As Double    '   Angle that should be between L1 and L2
    P0         As Integer    '     Common point of L1 and L2
    P1         As Integer    '     Other point on L1
    P2         As Integer    '     Other point on L2
    F          As Double    '       Muscle Force(strength)

    DynPhase   As Single
    DynAmp     As Single
    DynSpeed   As Single

    IsDynamic  As Boolean

    isNotBroken As Boolean

    FixedANG   As Boolean

End Type

Public Type tAABB
    Min        As tVector
    Max        As tVector
End Type




