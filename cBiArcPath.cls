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

Private NP As Long
Private BIARC() As cBiARC
Private NBia As Long

Private mTotalLength As Double

Private mClosed As Boolean


Friend Property Get Npoints() As Long
    Npoints = NP
End Property
Friend Property Get Length() As Double
    Length = mTotalLength
End Property

Friend Property Get Point(N As Long) As tVec2
    If N < 1 Then Exit Property
    If N > NP Then Exit Property
    If N = NP Then
        Point = BIARC(N - 1).Point2
    Else
        Point = BIARC(N).Point1
    End If
End Property
Friend Sub SetPoint(N As Long, Pos As tVec2, Optional autoTangent As Boolean = False)
    If N < NP Then BIARC(N).Point1 = Pos
    If N > 1 Then BIARC(N - 1).Point2 = Pos
    If autoTangent Then AutoSetTangents
End Sub
Friend Property Get TangDir(N As Long) As tVec2
    If N < 1 Then Exit Property
    If N > NP Then Exit Property
    If N = NP Then
        TangDir = BIARC(N - 1).TangentDir2
    Else
        TangDir = BIARC(N).TangentDir1
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
        NBia = NBia - 1: ReDim Preserve BIARC(NBia)
    End If
End Sub
Friend Sub AddPointAndControlPoint(P As tVec2, ControlP As tVec2)
    NP = NP + 1
    If NP = 1 Then
        FirstPoint = P
        FirstControlPoint = ControlP
    End If
    NBia = NP - 1
    If NBia > UBound(BIARC) Then
        ReDim Preserve BIARC(NBia)
        Set BIARC(NBia) = New cBiARC
        If NP = 2 Then
            BIARC(NBia).SetPointsAndControlPts FirstPoint, FirstControlPoint, _
                                               P, ControlP

        Else
            BIARC(NBia).SetPointsAndControlPts BIARC(NBia - 1).Point2, BIARC(NBia - 1).ControlPt2, _
                                               P, ControlP
        End If
    End If

    pvUpdatePathLength

End Sub
Friend Sub AddPointAndTangDirection(P As tVec2, TangDir As tVec2)
    AddPointAndControlPoint P, SUM2(P, TangDir)
End Sub

Friend Sub SetPointTangStart(PointN As Long, TangentStart As tVec2, Optional IsDirection As Boolean = False)
    If PointN > NP - 1 Then Exit Sub    'Last point hasn't TangStart
    If IsDirection Then
        BIARC(PointN).ControlPt1 = SUM2(BIARC(PointN).Point1, TangentStart)
    Else
        BIARC(PointN).ControlPt1 = TangentStart
    End If
    BIARC(PointN).CALC
    pvUpdatePathLength
End Sub
Friend Sub SetPointTangArrive(PointN As Long, TangentArrive As tVec2, Optional IsDirection As Boolean = False)
    If PointN < 2 Then Exit Sub                  '1st point hasnt arrive tangent
    If PointN > NP Then Exit Sub

    If IsDirection Then
        BIARC(PointN - 1).ControlPt2 = SUM2(BIARC(PointN - 1).Point2, TangentArrive)
    Else
        BIARC(PointN - 1).ControlPt2 = TangentArrive
    End If
    BIARC(PointN - 1).CALC
    pvUpdatePathLength
End Sub

Friend Sub SetPointTangent(PointN As Long, Tangent As tVec2, Optional IsDirection As Boolean = False)
    SetPointTangStart PointN, Tangent, IsDirection
    SetPointTangArrive PointN, Tangent, IsDirection
End Sub

Private Sub pvUpdatePathLength()
    Dim I&
    mTotalLength = 0
    For I = 1 To NBia
        mTotalLength = mTotalLength + BIARC(I).LengthTot
    Next
End Sub

Private Sub Class_Initialize()
    ReDim BIARC(0)
End Sub

Friend Sub DRAW(CC As cCairoContext, Color As Long, Alpha As Double, Width As Double, _
                Optional Lines As Boolean = False, Optional FullCircles As Boolean = False)
    Dim I&
    For I = 1 To NBia
        BIARC(I).DRAW CC, Color, Alpha, Width, Lines, FullCircles
    Next

End Sub

Friend Sub DrawOnlyCairoArcs(CC As cCairoContext)
    Dim I&
    'CC.MoveTo BIARC(1).Point1.x, BIARC(1).Point1.Y
    For I = 1 To NBia: BIARC(I).DrawArcsReadyToFillStroke CC: Next
End Sub

Friend Sub DrawFilled(CC As cCairoContext, Color As Long, Alpha As Double)
    CC.SetSourceColor Color, Alpha
    DrawOnlyCairoArcs CC
    CC.Fill
End Sub



Friend Function InterpolatedPointAt(ByVal T As Double) As tVec2
    Dim currT As Double
    Dim I&, J&
    If NP < 2 Then Exit Function
    '--- FIND which Biarc "T" belong to.   (I)
    T = T * mTotalLength
    I = 0
    Do
        I = I + 1
        currT = currT + BIARC(I).LengthTot
    Loop While T > currT And I < NBia
    '----------
    For J = 1 To I - 1
        T = T - BIARC(J).LengthTot
    Next
    T = T / BIARC(I).LengthTot

    InterpolatedPointAt = BIARC(I).InterpolatedPointAt(T)

End Function

Friend Sub Destroy()
    NP = 0
    NBia = 0
    ReDim BIARC(0)
    mClosed = False
End Sub

Friend Sub AutoSetTangents()
    Dim I As Long
    Dim Delta1 As tVec2, Delta2 As tVec2
    Dim NewTang As tVec2

    For I = 2 To NP - 1
        Delta1 = DIFF2(Point(I + 1), Point(I))
        Delta2 = DIFF2(Point(I), Point(I - 1))
'        NewTang = SUM2(Normalize2(Delta1), Normalize2(Delta2))
        NewTang = Normalize2(SUM2(Delta1, Delta2))

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



Friend Function CircleCollisionAndResponse(CenterC As tVec2, Vel As tVec2, ByVal Rad As Double, Optional Restitution As Double = 0.97) As tVec2
    Dim I         As Long
    Dim D         As Double
    Dim Circle1or2 As Long
    Dim rr        As Double
    Dim A         As Double
    Dim TANG      As tVec2
    Dim nTang     As tVec2
    Dim Vperp     As tVec2
    Dim vAxe      As tVec2
    For I = 1 To NBia
        With BIARC(I)
            D = .CircleDistance(CenterC, Rad, Circle1or2)
            If D < 0 Then
                '--- Displace
                If Circle1or2 = 1 Then           'Circle1
                    A = Atan2(CenterC.X - .Center1.X, CenterC.Y - .Center1.Y)
                    If .ClockWise1 Then rr = .Radius1 + Rad Else: rr = .Radius1 - Rad
                    TANG = Vec2(Cos(A) * rr, Sin(A) * rr)
                    CenterC = SUM2(.Center1, TANG)
                Else                             'Circle2
                    A = Atan2(CenterC.X - .Center2.X, CenterC.Y - .Center2.Y)
                    If .ClockWise2 Then rr = .Radius2 + Rad Else: rr = .Radius2 - Rad
                    TANG = Vec2(Cos(A) * rr, Sin(A) * rr)
                    CenterC = SUM2(.Center2, TANG)
                End If
                '---- Update VEL (bounce)
                vAxe = VectorProject(Vel, TANG)
                Vperp = DIFF2(Vel, vAxe)
                Vel = SUM2(Vperp, MUL2(vAxe, -Restitution))
                '-----------
                If Circle1or2 = 1 And Not (.ClockWise1) Then TANG = MUL2(TANG, -1)
                If Circle1or2 = 2 And Not (.ClockWise2) Then TANG = MUL2(TANG, -1)

                TANG = Normalize2(TANG)
                CircleCollisionAndResponse = Vec2(-TANG.Y, TANG.X)
                Exit For
            End If

        End With
    Next

End Function
