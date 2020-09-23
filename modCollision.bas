Attribute VB_Name = "modCollision"
Option Explicit


Public Type tBodyCollisionInfo

    BodyA      As Long    'New clsSoftBody
    BodyB      As Long    'New clsSoftBody

    bodyApm    As Long
    bodyBpmA   As Long
    bodyBpmB   As Long

    HitPt      As tVector
    EdgeD      As Single
    Normal     As tVector

    Penetration As Single

End Type

Public Type tMaterialPairs
    Friction   As Single    '= 0.3
    Elasticity As Single    '= 0.8f;
    Collide    As Boolean    '= true;
    '            mDefaultMatPair.CollisionFilter = new collisionFilter(this.defaultCollisionFilter);
End Type

Public MaterialPairs() As tMaterialPairs
Public MaterialCount As Long





Public BO()    As New clsSoftBody


Public CollisionList() As tBodyCollisionInfo

Public Const mPenetrationThreshold = 2000

Public mPenetrationCount As Long

Public Sub InitMatertials()
    Dim I      As Long
    Dim J      As Long

    MaterialCount = UBound(BO)
    ReDim MaterialPairs(1 To MaterialCount, 1 To MaterialCount)
    For I = 1 To MaterialCount

        For J = 1 To MaterialCount
            MaterialPairs(I, J).Friction = 0.7
            MaterialPairs(I, J).Elasticity = 0.8    '0.8
            MaterialPairs(I, J).Collide = True
        Next
    Next

End Sub

Public Sub ClearInfo(ByRef BCF As tBodyCollisionInfo)
    With BCF
        'BodyA = BodyB = Null
        .BodyA = 0
        .BodyB = 0
        .bodyApm = -1
        .bodyBpmA = -1
        .bodyBpmB = -1
        .HitPt.X = 0
        .HitPt.Y = 0
        .EdgeD = 0
    End With

End Sub


Public Sub infoListAdd(IL() As tBodyCollisionInfo, toAdd As tBodyCollisionInfo)
    ReDim Preserve IL(UBound(IL) + 1)
    With IL(UBound(IL))
        .BodyA = toAdd.BodyA
        .bodyApm = toAdd.bodyApm
        .BodyB = toAdd.BodyB
        .bodyBpmA = toAdd.bodyBpmA
        .bodyBpmB = toAdd.bodyBpmB
        .EdgeD = toAdd.EdgeD
        .HitPt = toAdd.HitPt
        .Normal = toAdd.Normal
        .Penetration = toAdd.Penetration



    End With

End Sub


Public Function AABBContains(AABB As tAABB, Pt As tVector) As Boolean
    AABBContains = (Pt.X >= AABB.Min.X) And (Pt.X <= AABB.Max.X) And (Pt.Y >= AABB.Min.Y) And (Pt.Y <= AABB.Max.Y)
End Function

Public Function AABBIntersects(AABB1 As tAABB, AABB2 As tAABB) As Boolean
    Dim OverlapX As Boolean
    Dim OverlapY As Boolean

    OverlapX = ((AABB1.Min.X <= AABB2.Max.X) And (AABB1.Max.X >= AABB2.Min.X))
    OverlapY = ((AABB1.Min.Y <= AABB2.Max.Y) And (AABB1.Max.Y >= AABB2.Min.Y))

    AABBIntersects = OverlapX And OverlapY


End Function




Public Sub BodyCollide(bA As Long, bB As Long, ByRef RetCollList() As tBodyCollisionInfo)
'COLLISION CHECKS / RESPONSE
    Dim bApmCount As Long
    Dim bBpmCount As Long
    Dim BoxA   As tAABB
    Dim BoxB   As tAABB
    Dim infoAway As tBodyCollisionInfo
    Dim infoSame As tBodyCollisionInfo
    Dim PrevPt As Long
    Dim NextPt As Long
    Dim Pre    As tVector
    Dim Nex    As tVector
    Dim fromPrev As tVector
    Dim toNext As tVector
    Dim ptNorm As tVector
    Dim closestAway As Single
    Dim closestSame As Single

    Dim Found  As Boolean

    Dim B1     As Long
    Dim B2     As Long

    Dim I      As Long
    Dim J      As Long

    Dim Pt     As tVector

    Dim HitPt  As tVector
    Dim Norm   As tVector
    Dim EdgeD  As Single
    Dim Pt1    As tVector
    Dim Pt2    As tVector

    Dim DistToA As Single
    Dim DistToB As Single
    Dim DOT    As Single
    Dim DIST   As Single


    bApmCount = BO(bA).NumS
    bBpmCount = BO(bB).NumS


    'AABB boxB = bB.getAABB();

    BoxB.Min.X = BO(bB).GetAABBMinX
    BoxB.Min.Y = BO(bB).GetAABBMinY
    BoxB.Max.X = BO(bB).GetAABBMaxX
    BoxB.Max.Y = BO(bB).GetAABBMaxY

    BoxA.Min.X = BO(bA).GetAABBMinX
    BoxA.Min.Y = BO(bA).GetAABBMinY
    BoxA.Max.X = BO(bA).GetAABBMaxX
    BoxA.Max.Y = BO(bA).GetAABBMaxY

    If Not AABBIntersects(BoxA, BoxB) Then Exit Sub


    '// check all PointMasses on bodyA for collision against bodyB.  if there is a collision, return detailed info.
    For I = 1 To bApmCount


        'Vector2 pt = bA.getPointMass(i).Position;
        Pt.X = BO(bA).GetPointPosX(I)
        Pt.Y = BO(bA).GetPointPosY(I)

        '// early out - if this point is outside the bounding box for bodyB, skip it!
        'if (!boxB.contains(ref pt))
        '    continue;
        If Not (AABBContains(BoxB, Pt)) Then GoTo ContinueIFor


        '// early out - if this point is not inside bodyB, skip it!
        'if (!bB.contains(ref pt))
        '    continue;
        If Not (BO(bB).Contains(Pt.X, Pt.Y)) Then GoTo ContinueIFor



        'int prevPt = (i>0) ? i-1 : bApmCount-1;
        'int nextPt = (i < bApmCount - 1) ? i + 1 : 0;

        PrevPt = IIf(I > 1, I - 1, bApmCount)
        NextPt = IIf(I < bApmCount, I + 1, 1)

        'Vector2 prev = bA.getPointMass(prevPt).Position;
        'Vector2 next = bA.getPointMass(nextPt).Position;

        Pre.X = BO(bA).GetPointPosX(PrevPt)
        Pre.Y = BO(bA).GetPointPosY(PrevPt)
        Nex.X = BO(bA).GetPointPosX(NextPt)
        Nex.Y = BO(bA).GetPointPosY(NextPt)


        '// now get the normal for this point. (NOT A UNIT VECTOR)
        'Vector2 fromPrev = new Vector2();
        'fromPrev.X = pt.X - prev.X;
        'fromPrev.Y = pt.Y - prev.Y;
        fromPrev.X = Pt.X - Pre.X
        fromPrev.Y = Pt.Y - Pre.Y

        'Vector2 toNext = new Vector2();
        'toNext.X = next.X - pt.X;
        'toNext.Y = next.Y - pt.Y;
        toNext.X = Nex.X - Pt.X
        toNext.Y = Nex.Y - Pt.Y


        'Vector2 ptNorm = new Vector2();
        'ptNorm.X = fromPrev.X + toNext.X;
        'ptNorm.Y = fromPrev.Y + toNext.Y;
        'VectorTools.makePerpendicular(ref ptNorm);

        ptNorm.X = fromPrev.X + toNext.X
        ptNorm.Y = fromPrev.Y + toNext.Y
        VectorPerp ptNorm



        '// this point is inside the other body.  now check if the edges on either side intersect with and edges on bodyB.
        'float closestAway = 100000.0f;
        'float closestSame = 100000.0f;
        closestAway = 999999
        closestSame = 999999

        'infoAway.clear();
        'infoAway.bodyA = bA;
        'infoAway.bodyApm = i;
        'infoAway.bodyB = bB;

        ClearInfo infoAway
        infoAway.BodyA = bA
        infoAway.bodyApm = I
        infoAway.BodyB = bB

        'infoSame.Clear();
        'infoSame.bodyA = bA;
        'infoSame.bodyApm = i;
        'infoSame.bodyB = bB;
        ClearInfo infoSame
        infoSame.BodyA = bA
        infoSame.bodyApm = I
        infoSame.BodyB = bB


        'bool found = false;
        'int b1 = 0;
        'int b2 = 1;
        Found = False
        B1 = 0
        B2 = 1

        For J = 1 To bBpmCount

            'Vector2 hitPt;
            'Vector2 norm;
            'float edgeD;
            'b1 = j;

            B1 = J

            If (J < bBpmCount) Then
                B2 = J + 1
            Else
                B2 = 1
            End If

            'Vector2 pt1 = bB.getPointMass(b1).Position;
            'Vector2 pt2 = bB.getPointMass(b2).Position;

            Pt1.X = BO(bB).GetPointPosX(B1)
            Pt1.Y = BO(bB).GetPointPosY(B1)
            Pt2.X = BO(bB).GetPointPosX(B2)
            Pt2.Y = BO(bB).GetPointPosY(B2)


            '// quick test of distance to each point on the edge, if both are greater than current mins, we can skip!
            'float distToA = ((pt1.X - pt.X) * (pt1.X - pt.X)) + ((pt1.Y - pt.Y) * (pt1.Y - pt.Y));
            'float distToB = ((pt2.X - pt.X) * (pt2.X - pt.X)) + ((pt2.Y - pt.Y) * (pt2.Y - pt.Y));

            DistToA = ((Pt1.X - Pt.X) * (Pt1.X - Pt.X)) + ((Pt1.Y - Pt.Y) * (Pt1.Y - Pt.Y))
            DistToB = ((Pt2.X - Pt.X) * (Pt2.X - Pt.X)) + ((Pt2.Y - Pt.Y) * (Pt2.Y - Pt.Y))



            'if ((distToA > closestAway) && (distToA > closestSame) && (distToB > closestAway) && (distToB > closestSame))
            '    continue;

            If ((DistToA > closestAway) And (DistToA > closestSame) And (DistToB > closestAway) And (DistToB > closestSame)) Then GoTo ContinueJFor



            '// test against this edge.
            'float dist = bB.getClosestPointOnEdgeSquared(pt, j, out hitPt, out norm, out edgeD);
            DIST = BO(bB).getClosestPointOnEdgeSquared(Pt.X, Pt.Y, J, HitPt.X, HitPt.Y, Norm.X, Norm.Y, EdgeD)



            '// only perform the check if the normal for this edge is facing AWAY from the point normal.
            'float dot;
            '//Vector2.Dot(ref ptNorm, ref edgeNorm, out dot);
            'Vector2.Dot(ref ptNorm, ref norm, out dot);

            DOT = VectorDot(ptNorm, Norm)


            If (DOT <= 0) Then

                If (DIST < closestAway) Then
                    closestAway = DIST
                    infoAway.bodyBpmA = B1
                    infoAway.bodyBpmB = B2
                    infoAway.EdgeD = EdgeD
                    infoAway.HitPt = HitPt
                    infoAway.Normal = Norm
                    infoAway.Penetration = DIST
                    Found = True
                End If

            Else

                If (DIST < closestSame) Then

                    closestSame = DIST
                    infoSame.bodyBpmA = B1
                    infoSame.bodyBpmB = B2
                    infoSame.EdgeD = EdgeD
                    infoSame.HitPt = HitPt
                    infoSame.Normal = Norm
                    infoSame.Penetration = DIST
                End If
            End If
ContinueJFor:
        Next J

        '// we've checked all edges on BodyB.  add the collision info to the stack.
        'if ((found) && (closestAway > mPenetrationThreshold) && (closestSame < closestAway))
        If ((Found) And (closestAway > mPenetrationThreshold) And (closestSame < closestAway)) Then

            'infoSame.penetration = (float)Math.Sqrt(infoSame.penetration);
            'if ( infoSame.bodyBpmA != -1 && infoSame.bodyBpmB != -1 )
            '   infoList.Add(infoSame);


            infoSame.Penetration = Sqr(infoSame.Penetration)

            If (infoSame.bodyBpmA <> -1 And infoSame.bodyBpmB <> -1) Then infoListAdd RetCollList(), infoSame

        Else

            'infoAway.penetration = (float)Math.Sqrt(infoAway.penetration);
            'if ( infoAway.bodyBpmA != -1 && infoAway.bodyBpmB != -1 )
            '    infoList.Add( infoAway );
            infoAway.Penetration = Sqr(infoAway.Penetration)
            If (infoAway.bodyBpmA <> -1 And infoAway.bodyBpmB <> -1) Then infoListAdd RetCollList(), infoAway

        End If

ContinueIFor:

    Next I


End Sub




'private void _handleCollisions()
'{

Public Sub HandleCollisions()
'            // handle all collisions!
    Dim I      As Long
    Dim Info   As tBodyCollisionInfo
    Dim A      As tPoint
    Dim B1     As tPoint
    Dim B2     As tPoint

    Dim bVel   As tVector
    Dim RelVel As tVector

    Dim RelDOT As Single

    Dim B1Inf  As Single
    Dim B2Inf  As Single

    Dim B2MassSum As Single
    Dim MassSum As Single
    Dim AMove  As Single
    Dim BMove  As Single
    Dim B1Move As Single
    Dim B2Move As Single

    Dim AInvMass As Single
    Dim BInvMass As Single

    Dim JDenom As Single
    Dim NumV   As tVector


    Dim Elas   As Single

    Dim JNumerator As Single

    Dim J      As Single

    Dim Tangent As tVector
    Dim Friction As Single
    Dim fNumerator As Single
    Dim continue As Boolean

    Dim F      As Single
    ' frmMAIN.Caption = UBound(CollisionList)

    'for (int i = 0; i < mCollisionList.Count; i++)
    For I = 1 To UBound(CollisionList)
        '                Stop



        Info = CollisionList(I)


        'PointMass A = info.bodyA.getPointMass(info.bodyApm);
        'PointMass B1 = info.bodyB.getPointMass(info.bodyBpmA);
        'PointMass B2 = info.bodyB.getPointMass(info.bodyBpmB);

        A.POS.X = BO(Info.BodyA).GetPointCpyPosX(Info.bodyApm)
        A.POS.Y = BO(Info.BodyA).GetPointCpyPosY(Info.bodyApm)
        A.VEL.X = BO(Info.BodyA).GetPointCpyVelX(Info.bodyApm)
        A.VEL.Y = BO(Info.BodyA).GetPointCpyVelY(Info.bodyApm)
        A.Mass = BO(Info.BodyA).GetPointMASS(Info.bodyApm)
        '-
        B1.POS.X = BO(Info.BodyB).GetPointCpyPosX(Info.bodyBpmA)
        B1.POS.Y = BO(Info.BodyB).GetPointCpyPosY(Info.bodyBpmA)
        B1.VEL.X = BO(Info.BodyB).GetPointCpyVelX(Info.bodyBpmA)
        B1.VEL.Y = BO(Info.BodyB).GetPointCpyVelY(Info.bodyBpmA)
        B1.Mass = BO(Info.BodyB).GetPointMASS(Info.bodyBpmA)

        '-
        B2.POS.X = BO(Info.BodyB).GetPointCpyPosX(Info.bodyBpmB)
        B2.POS.Y = BO(Info.BodyB).GetPointCpyPosY(Info.bodyBpmB)
        B2.VEL.X = BO(Info.BodyB).GetPointCpyVelX(Info.bodyBpmB)
        B2.VEL.Y = BO(Info.BodyB).GetPointCpyVelY(Info.bodyBpmB)
        B2.Mass = BO(Info.BodyB).GetPointMASS(Info.bodyBpmB)



        '// velocity changes as a result of collision.
        'Vector2 bVel = new Vector2();
        'bVel.X = (B1.Velocity.X + B2.Velocity.X) * 0.5f;
        'bVel.Y = (B1.Velocity.Y + B2.Velocity.Y) * 0.5f;
        bVel.X = (B1.VEL.X + B2.VEL.X) * 0.5
        bVel.Y = (B1.VEL.Y + B2.VEL.Y) * 0.5

        'Vector2 relVel = new Vector2();
        'relVel.X = A.Velocity.X - bVel.X;
        'relVel.Y = A.Velocity.Y - bVel.Y;

        RelVel.X = A.VEL.X - bVel.X
        RelVel.Y = A.VEL.Y - bVel.Y


        'float relDot;
        'Vector2.Dot(ref relVel, ref info.normal, out relDot);

        RelDOT = VectorDot(RelVel, Info.Normal)



        '// collision filter!
        'if (!mMaterialPairs[info.bodyA.Material, info.bodyB.Material].CollisionFilter(info.bodyA, info.bodyApm, info.bodyB, info.bodyBpmA, info.bodyBpmB, info.hitPt, relDot))
        '    continue;



        continue = False
        If (Info.Penetration > mPenetrationThreshold) Then

            '//Console.WriteLine("penetration above Penetration Threshold!!  penetration={0}  threshold={1} difference={2}",
            '//    info.penetration, mPenetrationThreshold, info.penetration-mPenetrationThreshold);
            MsgBox "penetration above Penetration Threshold!! " & Info.Penetration & "  " & mPenetrationThreshold & "  " & Info.Penetration - mPenetrationThreshold
            'mPenetrationCount++;
            'continue;
            mPenetrationCount = mPenetrationCount + 1
            continue = True

        End If
        If continue Then GoTo ContinueFor

        B1Inf = 1 - Info.EdgeD
        B2Inf = Info.EdgeD

        '                float b2MassSum = ((float.IsPositiveInfinity(B1.Mass)) || (float.IsPositiveInfinity(B2.Mass))) ? float.PositiveInfinity : (B1.Mass + B2.Mass);
        'float b2MassSum = (IsPositiveInfinity(B1Mass)) || (float.IsPositiveInfinity(B2.Mass))) ? float.PositiveInfinity : (
        'B2MassSum = Mass + Mass '(B1.Mass + B2.Mass);

        B2MassSum = IIf(B1.Mass = PositiveInfinity Or B2.Mass = PositiveInfinity, PositiveInfinity, B1.Mass + B2.Mass)


        MassSum = A.Mass + B2MassSum


        'If MassSum >= PositiveInfinity Then Stop


        If (A.Mass = PositiveInfinity) Then

            AMove = 0
            BMove = (Info.Penetration) + 0.1    ' 0.001

        ElseIf (B2MassSum = PositiveInfinity) Then

            AMove = (Info.Penetration) + 0.1    ' 0.001
            BMove = 0

        Else

            AMove = (Info.Penetration * (B2MassSum / MassSum))
            BMove = (Info.Penetration * (A.Mass / MassSum))
        End If


        ' AMove = (Info.Penetration * (B2MassSum / MassSum))
        ' BMove = (Info.Penetration * (A.Mass / MassSum))



        B1Move = BMove * B1Inf
        B2Move = BMove * B2Inf

        'float AinvMass = (float.IsPositiveInfinity(A.Mass)) ? 0f : 1f / A.Mass;
        'float BinvMass = (float.IsPositiveInfinity(b2MassSum)) ? 0f : 1f / b2MassSum;
        AInvMass = IIf(A.Mass = PositiveInfinity, 0, 1 / A.Mass)
        BInvMass = IIf(B2MassSum = PositiveInfinity, 0, 1 / B2MassSum)


        JDenom = AInvMass + BInvMass

        'float elas = 1f + mMaterialPairs[info.bodyA.Material, info.bodyB.Material].Elasticity;
        Elas = 1 + MaterialPairs(BO(Info.BodyA).Material, BO(Info.BodyB).Material).Elasticity



        NumV.X = RelVel.X * Elas
        NumV.Y = RelVel.Y * Elas

        'float jNumerator;
        'Vector2.Dot(ref numV, ref info.normal, out jNumerator);
        'jNumerator = -jNumerator;

        JNumerator = VectorDot(NumV, Info.Normal)
        JNumerator = -JNumerator


        J = JNumerator / JDenom




        'if (!float.IsPositiveInfinity(A.Mass))
        '{
        '    A.Position.X += info.normal.X * Amove;
        '    A.Position.Y += info.normal.Y * Amove;
        '}
        '
        '                if (!float.IsPositiveInfinity(B1.Mass))
        '                {
        '                    B1.Position.X -= info.normal.X * B1move;
        '                    B1.Position.Y -= info.normal.Y * B1move;
        '                }
        '
        '                if (!float.IsPositiveInfinity(B2.Mass))
        '                {
        '                    B2.Position.X -= info.normal.X * B2move;
        '                    B2.Position.Y -= info.normal.Y * B2move;
        '                }

        A.POS.X = A.POS.X + Info.Normal.X * AMove
        A.POS.Y = A.POS.Y + Info.Normal.Y * AMove

        B1.POS.X = B1.POS.X - Info.Normal.X * B1Move
        B1.POS.Y = B1.POS.Y - Info.Normal.Y * B1Move

        B2.POS.X = B2.POS.X - Info.Normal.X * B2Move
        B2.POS.Y = B2.POS.Y - Info.Normal.Y * B2Move


        '            GoTo skip

        BO(Info.BodyA).SetPointPosX(Info.bodyApm) = A.POS.X
        BO(Info.BodyA).SetPointPosY(Info.bodyApm) = A.POS.Y

        BO(Info.BodyB).SetPointPosX(Info.bodyBpmA) = B1.POS.X
        BO(Info.BodyB).SetPointPosY(Info.bodyBpmA) = B1.POS.Y

        BO(Info.BodyB).SetPointPosX(Info.bodyBpmB) = B2.POS.X
        BO(Info.BodyB).SetPointPosY(Info.bodyBpmB) = B2.POS.Y

skip:

        'Vector2 tangent = new Vector2();
        'VectorTools.getPerpendicular(ref info.normal, ref tangent);
        Tangent = Info.Normal
        VectorPerp Tangent


        'float friction = mMaterialPairs[info.bodyA.Material,info.bodyB.Material].Friction;
        Friction = MaterialPairs(BO(Info.BodyA).Material, BO(Info.BodyB).Material).Friction


        'Vector2.Dot(ref relVel, ref tangent, out fNumerator);
        'fNumerator *= friction;
        'float f = fNumerator / jDenom;
        fNumerator = VectorDot(RelVel, Tangent)
        fNumerator = fNumerator * Friction
        F = fNumerator / JDenom


        '// adjust velocity if relative velocity is moving toward each other.
        'if (relDot <= 0.0001f)

        'If RelDOT <= 0.0001 Then


        If RelDOT <= 0 Then
            'if (!float.IsPositiveInfinity(A.Mass))
            '{
            '    A.Velocity.X += (info.normal.X * (j / A.Mass)) - (tangent.X * (f / A.Mass));
            '    A.Velocity.Y += (info.normal.Y * (j / A.Mass)) - (tangent.Y * (f / A.Mass));
            '}
            'if (!float.IsPositiveInfinity(b2MassSum))
            '{
            '    B1.Velocity.X -= (info.normal.X * (j / b2MassSum) * b1inf) - (tangent.X * (f / b2MassSum) * b1inf);
            '    B1.Velocity.Y -= (info.normal.Y * (j / b2MassSum) * b1inf) - (tangent.Y * (f / b2MassSum) * b1inf);
            '}

            'if (!float.IsPositiveInfinity(b2MassSum))
            '{
            '    B2.Velocity.X -= (info.normal.X * (j / b2MassSum) * b2inf) - (tangent.X * (f / b2MassSum) * b2inf);
            '    B2.Velocity.Y -= (info.normal.Y * (j / b2MassSum) * b2inf) - (tangent.Y * (f / b2MassSum) * b2inf);
            '}

            A.VEL.X = A.VEL.X + (Info.Normal.X * (J / A.Mass)) - (Tangent.X * (F / A.Mass))
            A.VEL.Y = A.VEL.Y + (Info.Normal.Y * (J / A.Mass)) - (Tangent.Y * (F / A.Mass))

            B1.VEL.X = B1.VEL.X - (Info.Normal.X * (J / B2MassSum) * B1Inf) - (Tangent.X * (F / B2MassSum) * B1Inf)
            B1.VEL.Y = B1.VEL.Y - (Info.Normal.Y * (J / B2MassSum) * B1Inf) - (Tangent.Y * (F / B2MassSum) * B1Inf)

            B2.VEL.X = B2.VEL.X - (Info.Normal.X * (J / B2MassSum) * B2Inf) - (Tangent.X * (F / B2MassSum) * B2Inf)
            B2.VEL.Y = B2.VEL.Y - (Info.Normal.Y * (J / B2MassSum) * B2Inf) - (Tangent.Y * (F / B2MassSum) * B2Inf)



            BO(Info.BodyA).SetPointVelX(Info.bodyApm) = A.VEL.X
            BO(Info.BodyA).SetPointVelY(Info.bodyApm) = A.VEL.Y

            BO(Info.BodyB).SetPointVelX(Info.bodyBpmA) = B1.VEL.X
            BO(Info.BodyB).SetPointVelY(Info.bodyBpmA) = B1.VEL.Y

            BO(Info.BodyB).SetPointVelX(Info.bodyBpmB) = B2.VEL.X
            BO(Info.BodyB).SetPointVelY(Info.bodyBpmB) = B2.VEL.Y

        End If
ContinueFor:

    Next







    ReDim CollisionList(0)

    'mCollisionList.Clear();

End Sub
