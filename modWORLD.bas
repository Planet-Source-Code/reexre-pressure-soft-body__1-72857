Attribute VB_Name = "modWORLD"
Option Explicit

Public Const PI = 3.14159265358979

Public PicW    As Integer
Public PicH    As Integer


Public Const GY As Single = 100    ' 200 ' 150

Public Const Wall_Friction As Single = 0.9

Public Const AirResistence As Single = 0.996

Public PointToMove As Long

Public DT      As Single
Public PTime   As Single


Public Const PositiveInfinity = 99999




Public Sub ADDBODY(File As String, CposX, CposY)
    ReDim Preserve BO(UBound(BO) + 1)
    BO(UBound(BO)).Me_LOAD File, CposX, CposY
    BO(UBound(BO)).COLOR = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
End Sub
