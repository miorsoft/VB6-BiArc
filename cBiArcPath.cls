VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiArcPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'BIARC Path
'Author: reexe -miorsoft

'******************************************************************************

Option Explicit

Private FirstPoint As tVec2
Private FirstControlPoint As tVec2

Private NP        As Long
Private BiArc()   As cBiARC
Private NBia      As Long

Private mTotalLength As Double

Private mClosed   As Boolean

'----------- DRAW Image
Private SRF       As cCairoSurface
Private SourceSRF As cCairoSurface
Private CC        As cCairoContext
Private pW        As Long
Private pH        As Long
Private ImaScaleX#, ImaScaleY#
Attribute ImaScaleY.VB_VarUserMemId = 1073938444
Private InvScaleX#, InvScaleY#
Attribute InvScaleX.VB_VarUserMemId = 1073938446
Attribute InvScaleY.VB_VarUserMemId = 1073938446



Friend Property Get Npoints() As Long
    Npoints = NP
End Property
Friend Property Get Length() As Double
    Length = mTotalLength
End Property

Friend Property Get Point(ByVal N As Long) As tVec2
    If N < 1 Then Exit Property
    If N > NP Then Exit Property
    If N = NP Then
        Point = BiArc(N - 1).Point2
    Else
        Point = BiArc(N).Point1
    End If
End Property

'Friend Property Let Point(N As Long, pt As tVec2)
'    If N < 1 Then Exit Property
'    If N > NP Then Exit Property
'    If N = NP Then
'         BiArc(N - 1).Point2 = pt
'    Else
'         BiArc(N).Point1 = pt
'    End If
'End Property


Friend Sub SetPoint(ByVal N As Long, Pos As tVec2, Optional autoTangent As Boolean = False)


    If N < NP Then BiArc(N).Point1 = Pos
    If N > 1 Then BiArc(N - 1).Point2 = Pos
    If autoTangent Then AutoSetTangents
End Sub
Friend Property Get TangDir(ByVal N As Long) As tVec2
    If N < 1 Then Exit Property
    If N > NP Then Exit Property
    If N = NP Then
        TangDir = BiArc(N - 1).TangentDir2
    Else
        TangDir = BiArc(N).TangentDir1
    End If
End Property
Friend Property Get Closed() As Boolean
    Closed = mClosed
End Property
Friend Sub ClosePath()
    If mClosed = False Then
        mClosed = True
        AddPointAndTangDirection Point(1), TangDir(1)
    End If
End Sub
Friend Sub OpenPath()
    If mClosed = True Then
        mClosed = False
        NP = NP - 1
        NBia = NBia - 1: ReDim Preserve BiArc(NBia)
    End If
End Sub
Friend Sub AddPointAndControlPoint(P As tVec2, ControlP As tVec2)
    NP = NP + 1
    If NP = 1 Then
        FirstPoint = P
        FirstControlPoint = ControlP
    End If
    NBia = NP - 1
    If NBia > UBound(BiArc) Then
        ReDim Preserve BiArc(NBia)
        Set BiArc(NBia) = New cBiARC
        If NP = 2 Then
            BiArc(NBia).SetPointsAndControlPts FirstPoint, FirstControlPoint, _
                                               P, ControlP

        Else
            BiArc(NBia).SetPointsAndControlPts BiArc(NBia - 1).Point2, BiArc(NBia - 1).ControlPt2, _
                                               P, ControlP
        End If
    End If

    pvUpdatePathLength

End Sub
Friend Sub AddPointAndTangDirection(P As tVec2, TangDir As tVec2)
    AddPointAndControlPoint P, SUM2(P, TangDir)
End Sub

Friend Sub SetPointTangStart(ByVal PointN As Long, TangentStart As tVec2, Optional IsDirection As Boolean = False)
    If PointN > NP - 1 Then Exit Sub    'Last point hasn't TangStart
    If IsDirection Then
        BiArc(PointN).ControlPt1 = SUM2(BiArc(PointN).Point1, TangentStart)
    Else
        BiArc(PointN).ControlPt1 = TangentStart
    End If
    BiArc(PointN).CALC
    pvUpdatePathLength
End Sub
Friend Sub SetPointTangArrive(ByVal PointN As Long, TangentArrive As tVec2, Optional IsDirection As Boolean = False)
    If PointN < 2 Then Exit Sub   '1st point hasnt arrive tangent
    If PointN > NP Then Exit Sub

    If IsDirection Then
        BiArc(PointN - 1).ControlPt2 = SUM2(BiArc(PointN - 1).Point2, TangentArrive)
    Else
        BiArc(PointN - 1).ControlPt2 = TangentArrive
    End If
    BiArc(PointN - 1).CALC
    pvUpdatePathLength
End Sub

Friend Sub SetPointTangent(ByVal PointN As Long, Tangent As tVec2, Optional IsDirection As Boolean = False)
    SetPointTangStart PointN, Tangent, IsDirection
    SetPointTangArrive PointN, Tangent, IsDirection
End Sub

Private Sub pvUpdatePathLength()
    Dim I&
    mTotalLength = 0
    For I = 1 To NBia
        mTotalLength = mTotalLength + BiArc(I).LengthTot
    Next
End Sub

Private Sub Class_Initialize()
    ReDim BiArc(0)
End Sub

Friend Sub DRAW(CC As cCairoContext, ByVal Color As Long, ByVal Alpha As Double, ByVal Width As Double, _
                Optional Lines As Boolean = False, Optional FullCircles As Boolean = False)
    Dim I&
    For I = 1 To NBia
        BiArc(I).DRAW CC, Color, Alpha, Width, Lines, FullCircles
    Next

End Sub

Friend Sub DrawOnlyCairoArcs(CC As cCairoContext)
    Dim I&
    'CC.MoveTo BIARC(1).Point1.x, BIARC(1).Point1.Y
    For I = 1 To NBia: BiArc(I).DrawArcsReadyToFillStroke CC: Next
End Sub

Friend Sub DrawFilled(CC As cCairoContext, ByVal Color As Long, ByVal Alpha As Double)
    CC.SetSourceColor Color, Alpha
    DrawOnlyCairoArcs CC
    CC.Fill
End Sub



Friend Function InterpolatedPointAt(ByVal T As Double) As tVec2
    Dim currT     As Double
    Dim I&, J&
    If NP < 2 Then Exit Function
    '--- FIND which Biarc "T" belong to.   (I)
    T = T * mTotalLength
    I = 0
    Do
        I = I + 1
        currT = currT + BiArc(I).LengthTot
    Loop While T > currT And I < NBia
    '----------
    For J = 1 To I - 1
        T = T - BiArc(J).LengthTot
    Next
    T = T / BiArc(I).LengthTot

    InterpolatedPointAt = BiArc(I).InterpolatedPointAt(T)

End Function
Friend Function InterpolatedTANGAt(ByVal T As Double) As tVec2
    Dim currT     As Double
    Dim I&, J&
    If NP < 2 Then Exit Function
    '--- FIND which Biarc "T" belong to.   (I)
    T = T * mTotalLength
    I = 0
    Do
        I = I + 1
        currT = currT + BiArc(I).LengthTot
    Loop While T > currT And I < NBia
    '----------
    For J = 1 To I - 1
        T = T - BiArc(J).LengthTot
    Next
    T = T / BiArc(I).LengthTot

    InterpolatedTANGAt = BiArc(I).InterpolatedTANGAt(T)

End Function

Friend Sub Destroy()
    NP = 0
    NBia = 0
    ReDim BiArc(0)
    mClosed = False
End Sub

Friend Sub AutoSetTangents()
    Dim I         As Long
    Dim Delta1 As tVec2, Delta2 As tVec2
    Dim NewTang   As tVec2



    For I = 2 To NP - 1

        ' NewTang = Normalize2(DIFF2(Point(I + 1), Point(I - 1)))

        Delta1 = DIFF2(Point(I + 1), Point(I))
        Delta2 = DIFF2(Point(I), Point(I - 1))
        NewTang = SUM2( _
                  MUL2(Delta1, 1 / Length2(Delta2)), _
                  MUL2(Delta2, 1 / Length2(Delta1)))

        SetPointTangent I, NewTang, True
    Next
    If mClosed Then
        SetPointTangent NP, DIFF2(Point(2), Point(NP - 1)), True
        SetPointTangent 1, TangDir(NP), True
    Else
        SetPointTangent NP, DIFF2(Point(NP), Point(NP - 1)), True
        SetPointTangent 1, DIFF2(Point(2), Point(1)), True
    End If
    

End Sub



Friend Function CircleCollisionAndResponseUsingSide(CircleCenter As tVec2, Vel As tVec2, ByVal CircleRadius As Double, Optional Restitution As Double = 0.97, Optional LeftSide As Boolean = True) As tVec2
    Dim I         As Long
    Dim d         As Double
    Dim outCircle1or2 As Long
    Dim rr        As Double
    Dim A         As Double
    Dim NORMAL    As tVec2
    Dim nTang     As tVec2
    Dim VelTangent As tVec2
    Dim VelNormal As tVec2
    Dim tRAD      As Double
    tRAD = CircleRadius
    For I = 1 To NBia
        With BiArc(I)
            d = .CircleDistance(CircleCenter, CircleRadius, outCircle1or2, LeftSide)
            If d < 0 Then
                If LeftSide Then tRAD = CircleRadius Else: tRAD = -CircleRadius    ' USED
                If outCircle1or2 = 1 Then    'Circle1
                    NORMAL = Vec2(CircleCenter.X - .Center1.X, CircleCenter.Y - .Center1.Y)
                    If .ClockWise1 Then rr = .Radius1 + tRAD Else: rr = .Radius1 - tRAD
                    NORMAL = MUL2(Normalize2(NORMAL), rr)
                    CircleCenter = SUM2(.Center1, NORMAL)
                Else              'Circle2
                    NORMAL = Vec2(CircleCenter.X - .Center2.X, CircleCenter.Y - .Center2.Y)
                    If .ClockWise2 Then rr = .Radius2 + tRAD Else: rr = .Radius2 - tRAD
                    NORMAL = MUL2(Normalize2(NORMAL), rr)
                    CircleCenter = SUM2(.Center2, NORMAL)
                End If
                '---- Update VEL (bounce)
                VelNormal = VectorProject(Vel, NORMAL)
                VelTangent = DIFF2(Vel, VelNormal)
                Vel = SUM2(VelTangent, MUL2(VelNormal, -Restitution))
                '-----------
                '                If outCircle1or2 = 1 And Not (.ClockWise1) Then NORMAL = MUL2(NORMAL, -1)
                '                If outCircle1or2 = 2 And Not (.ClockWise2) Then NORMAL = MUL2(NORMAL, -1)

                NORMAL = Normalize2(NORMAL)
                CircleCollisionAndResponseUsingSide = Vec2(-NORMAL.Y, NORMAL.X)
                Exit For
            End If

        End With
    Next

End Function

Friend Function CircleCollisionAndResponse(CircleCenter As tVec2, Vel As tVec2, ByVal CircleRadius As Double, Optional Restitution As Double = 0.97) As tVec2
    Dim I         As Long
    Dim d         As Double
    Dim outCircle1or2 As Long
    Dim rr        As Double
    Dim A         As Double
    Dim NORMAL    As tVec2
    Dim nTang     As tVec2
    Dim VelTangent As tVec2
    Dim VelNormal As tVec2
    Dim cDist     As Double
    For I = 1 To NBia
        With BiArc(I)
            d = .PointDistance(CircleCenter, outCircle1or2)    '<<<.......
            ' D>0 means the point is betweens a Arc Angle
            If d > 0 Then d = d - CircleRadius
            If d < 0 Then
                '----------- BALL center displacement so that it's collinerar to curve
                If outCircle1or2 = 1 Then    'Circle1
                    NORMAL = Vec2(CircleCenter.X - .Center1.X, CircleCenter.Y - .Center1.Y)
                    cDist = Length2(DIFF2(.Center1, CircleCenter))
                    If cDist < .Radius1 Then CircleRadius = -CircleRadius
                    NORMAL = MUL2(Normalize2(NORMAL), CircleRadius - (cDist - .Radius1))
                    CircleCenter = SUM2(CircleCenter, NORMAL)
                Else              'Circle2
                    NORMAL = Vec2(CircleCenter.X - .Center2.X, CircleCenter.Y - .Center2.Y)
                    cDist = Length2(DIFF2(.Center2, CircleCenter))
                    If cDist < .Radius2 Then CircleRadius = -CircleRadius
                    NORMAL = MUL2(Normalize2(NORMAL), CircleRadius - (cDist - .Radius2))
                    CircleCenter = SUM2(CircleCenter, NORMAL)
                End If
                '---- Update VEL (bounce)
                NORMAL = Normalize2(NORMAL)
                VelNormal = VectorProjectN(Vel, NORMAL)
                VelTangent = DIFF2(Vel, VelNormal)
                Vel = SUM2(VelTangent, MUL2(VelNormal, -Restitution))
                '-----------
                CircleCollisionAndResponse = Vec2(-NORMAL.Y, NORMAL.X)
                Exit For
            End If
        End With
    Next
End Function
Friend Function CircleCollisionAndResponse2(BallCenter As tVec2, Vel As tVec2, ByVal BallRadius As Double, ByRef InOutAngVel As Double, Optional Restitution As Double = 0.97) As tVec2
    Dim I         As Long
    Dim d         As Double
    Dim outCircle1or2 As Long
    Dim rr        As Double
    Dim A         As Double
    Dim NORMAL    As tVec2
    Dim nTang     As tVec2
    Dim VelTangent As tVec2
    Dim TVelTangent As tVec2
    Dim VelNormal As tVec2
    Dim cDist     As Double
    Dim VelTDot   As Double
    Dim S         As Double
    S = 1
    For I = 1 To NBia
        With BiArc(I)
            d = .PointDistance(BallCenter, outCircle1or2)    '<<<.......
            ' D>0 means the point is betweens a Arc Angle
            If d > 0 Then d = d - BallRadius
            If d < 0 Then
           
                '----------- BALL center displacement so that it's collinerar to curve
                If outCircle1or2 = 1 Then    'Circle1
                    NORMAL = Vec2(BallCenter.X - .Center1.X, BallCenter.Y - .Center1.Y)
                    cDist = Length2(DIFF2(.Center1, BallCenter))
                    If cDist < .Radius1 Then BallRadius = -BallRadius: S = -S
                    NORMAL = MUL2(Normalize2(NORMAL), BallRadius - (cDist - .Radius1))
                    BallCenter = SUM2(BallCenter, NORMAL)
                Else              'Circle2
                    NORMAL = Vec2(BallCenter.X - .Center2.X, BallCenter.Y - .Center2.Y)
                    cDist = Length2(DIFF2(.Center2, BallCenter))
                    If cDist < .Radius2 Then BallRadius = -BallRadius: S = -S
                    NORMAL = MUL2(Normalize2(NORMAL), BallRadius - (cDist - .Radius2))
                    BallCenter = SUM2(BallCenter, NORMAL)
                End If
                '---- Update VEL (bounce)
                NORMAL = Normalize2(NORMAL)
                VelNormal = VectorProjectN(Vel, NORMAL)
                VelTangent = DIFF2(Vel, VelNormal)
                TVelTangent = VelTangent

                VelTDot = DOT2(TVelTangent, Rotate90(NORMAL))

                ' ?????????????????
                'VelTangent = SUM2(MUL2(VelTangent, 0.9), MUL2(Rotate90(NORMAL), -0.01 * S * VelTDot * BallRadius * InOutAngVel))

                '-----------

                InOutAngVel = InOutAngVel * 0.5 + S * VelTDot / BallRadius


                Vel = SUM2(VelTangent, MUL2(VelNormal, -Restitution))
                CircleCollisionAndResponse2 = Rotate90(NORMAL)
                Exit For
            End If
        End With
    Next
End Function


Friend Sub SetupImage(SourceImageFileName As String, TargetSrf As cCairoSurface, Optional ScaleX As Double = 1, Optional ScaleY As Double = 1)
    Set SRF = TargetSrf
    Set CC = SRF.CreateContext

    ImaScaleX = ScaleX
    ImaScaleY = ScaleY
    InvScaleX = 1 / ImaScaleX
    InvScaleY = 1 / ImaScaleY

    Set SourceSRF = Cairo.ImageList.AddImage("sourceImage", SourceImageFileName)

    pW = SRF.Width - 1
    pH = SRF.Height - 1

End Sub

Friend Sub SetupImageScale(ByVal ScaleX As Double, ByVal ScaleY As Double)

    ImaScaleX = ScaleX
    ImaScaleY = ScaleY
    InvScaleX = 1 / ImaScaleX
    InvScaleY = 1 / ImaScaleY

End Sub
Friend Sub DrawIMAGE(ByVal ClearScreen As Boolean, Optional Length As Double = 0)
    '    Dim ArrSource() As Long
    '    Dim ArrSrf2() As Long
    Dim ArrSource() As Byte
    Dim ArrSrf2() As Byte

    Dim H2&, H22#
    Dim dx#, dy#
    Dim X&, Y&
    Dim XF&, XT&
    Dim YF&, YT&
    Dim d#
    Dim w&, h&
    Dim A#
    Dim GetX#, GetY#
    Dim bA        As tVec2
    Dim bB        As tVec2
    Dim I         As Long
    Dim SgnClockWise1 As Double
    Dim SgnClockWise2 As Double
    Dim PrevLength#
    Dim X4        As Long
    Dim GetX4     As Long
    Dim dmin1#, dmax1#
    Dim dmin2#, dmax2#
    Dim Alpha     As Double
    Dim Klen#
    If ClearScreen Then CC.SetSourceRGB 0.33, 0.33, 0.33: CC.Paint

    '    SourceSRF.BindToArrayLong ArrSource
    '    Srf.BindToArrayLong ArrSrf2
    SourceSRF.BindToArray ArrSource
    SRF.BindToArray ArrSrf2

    w = UBound(ArrSource, 1) \ 4
    h = UBound(ArrSource, 2)

    H2 = h / 2
    H22 = H2 * H2

    If Length Then Klen = Length / mTotalLength Else: Klen = 1

    For I = 1 To NBia
        With BiArc(I)
    
        
            .GetBounds bA, bB

            XF = bA.X - H2 * ImaScaleY
            YF = bA.Y - H2 * ImaScaleY
            XT = bB.X + H2 * ImaScaleY
            YT = bB.Y + H2 * ImaScaleY

            If .ClockWise1 Then SgnClockWise1 = 1 Else: SgnClockWise1 = -1
            If .ClockWise2 Then SgnClockWise2 = 1 Else: SgnClockWise2 = -1

            'xF = min(.Center1.X - .Radius1 * SgnClockWise1, .Center2.X - .Radius2 * SgnClockWise2) - H2 * ImaScaleY
            'yF = min(.Center1.Y - .Radius1 * SgnClockWise1, .Center2.Y - .Radius2 * SgnClockWise2) - H2 * ImaScaleY
            'xT = max(.Center1.X + .Radius1 * SgnClockWise1, .Center2.X + .Radius2 * SgnClockWise2) + H2 * ImaScaleY
            'yT = max(.Center1.Y + .Radius1 * SgnClockWise1, .Center2.Y + .Radius2 * SgnClockWise2) + H2 * ImaScaleY

            If XF < 0 Then XF = 0
            If YF < 0 Then YF = 0
            If XT > pW Then XT = pW
            If YT > pH Then YT = pH

            ''Debug
            'fImage.Caption = xF & " - " & xT & "   " & yF & " - " & yT

            dmin1 = .Radius1 - H2 * ImaScaleY
            dmin1 = Sgn(dmin1) * dmin1 * dmin1
            dmax1 = .Radius1 + H2 * ImaScaleY
            dmax1 = dmax1 * dmax1

            dmin2 = .Radius2 - H2 * ImaScaleY
            dmin2 = Sgn(dmin2) * dmin2 * dmin2
            dmax2 = .Radius2 + H2 * ImaScaleY
            dmax2 = dmax2 * dmax2


            For X = XF To XT
                X4 = X * 4
                For Y = YF To YT
                    dx = X - .Center1.X
                    dy = Y - .Center1.Y
                    d = dx * dx + dy * dy
                    If d > dmin1 And d < dmax1 Then
                        A = Atan2(dx, dy)
                        If IsAngBetween(A, .A11, .A12) Then
                            d = Sqr(d) - .Radius1
                            GetY = H2 - d * SgnClockWise1 * InvScaleY
                            If .ClockWise1 Then
                                GetX = AngleDIFF(A, .A11) * .Radius1
                            Else
                                GetX = -AngleDIFF(A, .A12) * .Radius1
                            End If
                            GetX = GetX + PrevLength
                            GetX = (GetX * InvScaleX * Klen) Mod w
                            GetY = GetY Mod h
                            ' ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                            GetX4 = GetX * 4&
                            '                            ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)

                            Alpha = ArrSource(GetX4 + 3&, GetY)
                            If Alpha Then
                            Alpha = Alpha * 3.92156862745098E-03
                                ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                                ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                            End If
                        End If
                    End If
                    dx = X - .Center2.X
                    dy = Y - .Center2.Y
                    d = dx * dx + dy * dy
                    If d > dmin2 And d < dmax2 Then
                        A = Atan2(dx, dy)
                        If IsAngBetween(A, .A21, .A22) Then
                        
                            d = Sqr(d) - .Radius2
                            GetY = H2 - d * SgnClockWise2 * InvScaleY

                            If .ClockWise2 Then
                                GetX = AngleDIFF(A, .A21) * .Radius2
                            Else
                                GetX = -AngleDIFF(A, .A22) * .Radius2
                            End If
                            GetX = GetX + .LengthArc1
                            GetX = GetX + PrevLength
                            GetX = (GetX * InvScaleX * Klen) Mod w
                            GetY = GetY Mod h
                            'ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                            GetX4 = GetX * 4&

                            '                            ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)

                            Alpha = ArrSource(GetX4 + 3&, GetY)
                            If Alpha Then
                            Alpha = Alpha * 3.92156862745098E-03
                            
                                ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                                ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                            End If
                        End If
                    End If

                Next
            Next

            PrevLength = PrevLength + .LengthTot

        End With
    Next                          'biarc

    '
    '    SourceSRF.ReleaseArrayLong ArrSource
    '    Srf.ReleaseArrayLong ArrSrf2

    SourceSRF.ReleaseArray ArrSource
    SRF.ReleaseArray ArrSrf2


End Sub



Friend Sub DrawIMAGEAA(ByVal ClearScreen As Boolean, Optional Length As Double = 0)

    ' TESTING Source Pixel AA  (not so good)

    '    Dim ArrSource() As Long
    '    Dim ArrSrf2() As Long
    Dim ArrSource() As Byte
    Dim ArrSrf2() As Byte

    Dim H2&, H22#
    Dim dx#, dy#
    Dim X&, Y&
    Dim XF&, XT&
    Dim YF&, YT&
    Dim d#
    Dim w&, h&
    Dim A#
    Dim GetX#, GetY#
    Dim bA        As tVec2
    Dim bB        As tVec2
    Dim I         As Long
    Dim SgnClockWise1 As Double
    Dim SgnClockWise2 As Double
    Dim PrevLength#
    Dim X4        As Long
    Dim GetX4     As Long
    Dim dmin1#, dmax1#
    Dim dmin2#, dmax2#
    Dim Alpha     As Double
    Dim Klen#
    Dim U#, V#
    If ClearScreen Then CC.SetSourceRGB 0.33, 0.33, 0.33: CC.Paint

    '    SourceSRF.BindToArrayLong ArrSource
    '    Srf.BindToArrayLong ArrSrf2
    SourceSRF.BindToArray ArrSource
    SRF.BindToArray ArrSrf2

    w = UBound(ArrSource, 1) \ 4
    h = UBound(ArrSource, 2)

    H2 = h / 2
    H22 = H2 * H2

    If Length Then Klen = Length / mTotalLength Else: Klen = 1


    Dim c00       As tRGB
    Dim c10       As tRGB
    Dim c01       As tRGB
    Dim c11       As tRGB
    Dim CI        As tRGB
    Dim GX#, GY#

    For I = 1 To NBia
        With BiArc(I)
            .GetBounds bA, bB

            XF = bA.X - H2 * ImaScaleY
            YF = bA.Y - H2 * ImaScaleY
            XT = bB.X + H2 * ImaScaleY
            YT = bB.Y + H2 * ImaScaleY

            If XF < 0 Then XF = 0
            If YF < 0 Then YF = 0
            If XT > pW Then XT = pW
            If YT > pH Then YT = pH

            dmin1 = .Radius1 - H2 * ImaScaleY
            dmin1 = Sgn(dmin1) * dmin1 * dmin1
            dmax1 = .Radius1 + H2 * ImaScaleY
            dmax1 = dmax1 * dmax1

            dmin2 = .Radius2 - H2 * ImaScaleY
            dmin2 = Sgn(dmin2) * dmin2 * dmin2
            dmax2 = .Radius2 + H2 * ImaScaleY
            dmax2 = dmax2 * dmax2

            If .ClockWise1 Then SgnClockWise1 = 1 Else: SgnClockWise1 = -1
            If .ClockWise2 Then SgnClockWise2 = 1 Else: SgnClockWise2 = -1


            For X = XF To XT
                X4 = X * 4
                For Y = YF To YT
                    dx = X - .Center1.X
                    dy = Y - .Center1.Y
                    d = dx * dx + dy * dy
                    If d > dmin1 And d < dmax1 Then
                        A = Atan2(dx, dy)
                        If IsAngBetween(A, .A11, .A12) Then
                            d = Sqr(d) - .Radius1
                            GetY = H2 - d * SgnClockWise1 * InvScaleY
                            If .ClockWise1 Then
                                GetX = AngleDIFF(A, .A11) * .Radius1
                            Else
                                GetX = -AngleDIFF(A, .A12) * .Radius1
                            End If
                            GetX = GetX + PrevLength
                            GetX = (GetX * InvScaleX * Klen)
                            GX = GetX
                            GetX = GetX Mod w
                            GY = GetY
                            GetY = GetY Mod h
                            ' ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                            GetX4 = GetX * 4&
                            '                            ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)



                            Alpha = ArrSource(GetX4 + 3&, GetY) * 3.92156862745098E-03
                            If Alpha Then
                                '    ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                     ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)

                                c00 = Col3(ArrSource(GetX4 + 2, GetY), ArrSource(GetX4 + 1, GetY), ArrSource(GetX4, GetY))
                                c10 = Col3(ArrSource(GetX4 + 2 + 4, GetY), ArrSource(GetX4 + 1 + 4, GetY), ArrSource(GetX4 + 4, GetY))
                                c01 = Col3(ArrSource(GetX4 + 2, GetY + 1), ArrSource(GetX4 + 1, GetY + 1), ArrSource(GetX4, GetY + 1))
                                c11 = Col3(ArrSource(GetX4 + 2 + 4, GetY + 1), ArrSource(GetX4 + 1 + 4, GetY + 1), ArrSource(GetX4 + 4, GetY + 1))


                                U = (GX - Int(GX))
                                V = (GY - Int(GY))

                                CI = Col3BilinearInterpolation(c00, c10, c01, c11, U, V)

                                ColorMixByAlpha CI.X * 1, CI.Y * 1, CI.z * 1, Alpha, _
                                                ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)



                            End If
                        End If
                    End If
                    dx = X - .Center2.X
                    dy = Y - .Center2.Y
                    d = dx * dx + dy * dy
                    If d > dmin2 And d < dmax2 Then
                        A = Atan2(dx, dy)
                        If IsAngBetween(A, .A21, .A22) Then
                            d = Sqr(d) - .Radius2
                            GetY = H2 - d * SgnClockWise2 * InvScaleY

                            If .ClockWise2 Then
                                GetX = AngleDIFF(A, .A21) * .Radius2
                            Else
                                GetX = -AngleDIFF(A, .A22) * .Radius2
                            End If
                            GetX = GetX + .LengthArc1
                            GetX = GetX + PrevLength
                            GetX = (GetX * InvScaleX * Klen)
                            GX = GetX
                            GetX = GetX Mod w
                            GY = GetY
                            GetY = GetY Mod h
                            'ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                            GetX4 = GetX * 4&

                            '                            ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)

                            Alpha = ArrSource(GetX4 + 3&, GetY) * 3.92156862745098E-03
                            If Alpha Then
                                '                                ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                                                 '                                                ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                                c00 = Col3(ArrSource(GetX4 + 2, GetY), ArrSource(GetX4 + 1, GetY), ArrSource(GetX4, GetY))
                                c10 = Col3(ArrSource(GetX4 + 2 + 4, GetY), ArrSource(GetX4 + 1 + 4, GetY), ArrSource(GetX4 + 4, GetY))
                                c01 = Col3(ArrSource(GetX4 + 2, GetY + 1), ArrSource(GetX4 + 1, GetY + 1), ArrSource(GetX4, GetY + 1))
                                c11 = Col3(ArrSource(GetX4 + 2 + 4, GetY + 1), ArrSource(GetX4 + 1 + 4, GetY + 1), ArrSource(GetX4 + 4, GetY + 1))



                                U = (GX - Int(GX))
                                V = (GY - Int(GY))

                                CI = Col3BilinearInterpolation(c00, c10, c01, c11, U, V)

                                ColorMixByAlpha CI.X * 1, CI.Y * 1, CI.z * 1, Alpha, _
                                                ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                            End If
                        End If
                    End If

                Next
            Next

            PrevLength = PrevLength + .LengthTot

        End With
    Next                          'biarc

    '
    '    SourceSRF.ReleaseArrayLong ArrSource
    '    Srf.ReleaseArrayLong ArrSrf2

    SourceSRF.ReleaseArray ArrSource
    SRF.ReleaseArray ArrSrf2




End Sub
Private Sub ColorMixByAlpha(ByVal iR As Byte, ByVal iG As Byte, ByVal iB As Byte, ByVal iAlpha As Double, outR As Byte, outG As Byte, outB As Byte)
    outR = iR * (iAlpha) + outR * (1# - iAlpha)
    outG = iG * (iAlpha) + outG * (1# - iAlpha)
    outB = iB * (iAlpha) + outB * (1# - iAlpha)
End Sub




Friend Sub DrawImageRC()
    Dim currT     As Double
    Dim I&, J&
    Dim fetta     As cCairoSurface
    Dim h         As Long
    Dim w         As Long


    Dim fettaOffset#
    Dim X1#, X3#
    Dim Y1#, Y3#
    Dim X2#, X4#
    Dim Y2#, Y4#
    Dim A1#
    Dim A2#
    Dim tmpV      As tVec2
    Dim wA&
    Dim T#, d#
    Dim AO#
    Dim S#
    Dim H2#
    Dim CosA1#
    Dim SinA1#
    Dim CosA2#
    Dim SinA2#

    If NP < 2 Then Exit Sub
    '--- FIND which Biarc "T" belong to.   (I)

    Const PASSO   As Double = 25

    w = SourceSRF.Width
    h = SourceSRF.Height
    H2 = h * 0.5
    CC.SetSourceRGB 0.5, 0.5, 0.5: CC.Paint


    Set fetta = Cairo.CreateSurface(PASSO, h)



    '    For T = 0 To mTotalLength Step passo
    For d = 0 To mTotalLength Step PASSO
        fettaOffset = d

        While fettaOffset > (w - PASSO): fettaOffset = fettaOffset - w: Wend


        I = 0
        currT = 0
        Do
            I = I + 1
            currT = currT + BiArc(I).LengthTot
        Loop While d > currT And I < NBia
        '----------
        T = d
        For J = 1 To I - 1
            T = T - BiArc(J).LengthTot
        Next
        AO = T
        T = T / BiArc(I).LengthTot


        With BiArc(I)
            Set fetta = Cairo.CreateSurface(PASSO, h)
            With fetta.CreateContext
                '                .SetSourceRGB 0.5, 0.5, 0.5: .Paint
                .RenderSurfaceContent SourceSRF, -fettaOffset, 0
                '            fetta.CreateContext.RenderSurfaceContent SourceSRF, 0, 0
            End With

            tmpV = .InterpolatedPointAt(T, wA)

            If wA = 1 Then
                A1 = Atan2(tmpV.X - .Center1.X, tmpV.Y - .Center1.Y)
                If .ClockWise1 Then S = 1 Else: S = -1
                A2 = A1 + S * (PASSO / .Radius1)

            Else
                A1 = Atan2(tmpV.X - .Center2.X, tmpV.Y - .Center2.Y)
                If .ClockWise2 Then S = 1 Else: S = -1
                A2 = A1 + S * (PASSO / .Radius2)
            End If

            CosA1 = Cos(A1): SinA1 = Sin(A1)
            CosA2 = Cos(A2): SinA2 = Sin(A2)

            If wA = 1 Then

                X1 = .Center1.X + (.Radius1 + H2 * S) * CosA1
                Y1 = .Center1.Y + (.Radius1 + H2 * S) * SinA1
                X2 = .Center1.X + (.Radius1 + H2 * S) * CosA2
                Y2 = .Center1.Y + (.Radius1 + H2 * S) * SinA2
                X3 = .Center1.X + (.Radius1 - H2 * S) * CosA2
                Y3 = .Center1.Y + (.Radius1 - H2 * S) * SinA2
                X4 = .Center1.X + (.Radius1 - H2 * S) * CosA1
                Y4 = .Center1.Y + (.Radius1 - H2 * S) * SinA1
            Else


                X1 = .Center2.X + (.Radius2 + H2 * S) * CosA1
                Y1 = .Center2.Y + (.Radius2 + H2 * S) * SinA1
                X2 = .Center2.X + (.Radius2 + H2 * S) * CosA2
                Y2 = .Center2.Y + (.Radius2 + H2 * S) * SinA2
                X3 = .Center2.X + (.Radius2 - H2 * S) * CosA2
                Y3 = .Center2.Y + (.Radius2 - H2 * S) * SinA2
                X4 = .Center2.X + (.Radius2 - H2 * S) * CosA1
                Y4 = .Center2.Y + (.Radius2 - H2 * S) * SinA1
            End If

        End With

        CC.RenderSurfaceWithinCorners fetta, X1, Y1, X2, Y2, X3, Y3, X4, Y4

        Debug.Print T, fettaOffset, AO
        DoEvents

    Next


End Sub


Friend Sub Offset(Amount#)
Dim I&


End Sub
