Attribute VB_Name = "modVector"
Option Explicit

Public Function RotatedVector(V As tVector, angleRadians As Single) As tVector
    Dim C      As Single
    Dim S      As Single

    C = Cos(angleRadians)
    S = Sin(angleRadians)
    RotatedVector.X = (C * V.X) - (S * V.Y)
    RotatedVector.Y = (C * V.Y) + (S * V.X)

End Function

Public Function Vector(XX, YY) As tVector
    Vector.X = XX
    Vector.Y = YY

End Function

Public Function VectorCross3(V1 As tVector3, V2 As tVector3) As tVector3

    VectorCross3.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
    VectorCross3.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
    VectorCross3.Z = (V1.X * V2.Y) - (V1.Y * V2.X)

End Function

Public Function VectorDist(V1 As tVector, V2 As tVector) As Single
    Dim dX     As Single
    Dim dY     As Single

    dX = V1.X - V2.X
    dY = V1.Y - V2.Y
    VectorDist = Sqr(dX * dX + dY * dY)


End Function

Public Function VectorDot(V1 As tVector, V2 As tVector) As Single

    VectorDot = (V1.X * V2.X) + _
                (V1.Y * V2.Y)


End Function

Public Function VectorNormalize(V As tVector) As tVector
    Dim Le     As Single
    Le = Sqr(V.X * V.X + V.Y * V.Y)
    VectorNormalize.X = V.X / Le
    VectorNormalize.Y = V.Y / Le
End Function

Public Function VectorPerp(ByRef V As tVector)
    Dim X      As Single
    X = V.X
    V.X = -V.Y
    V.Y = X
End Function

Public Function VectorSquaredDist(V1 As tVector, V2 As tVector) As Single
    Dim dX     As Single
    Dim dY     As Single

    dX = V1.X - V2.X
    dY = V1.Y - V2.Y
    VectorSquaredDist = (dX * dX + dY * dY)


End Function

