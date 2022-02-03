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
Private FirstPointTang As tVec2

Private NP        As Long
Private BIARC()   As cBiARC
Private NBia      As Long

Private mTotalLength As Double

Private mClosed   As Boolean


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
Friend Sub AddPointAndTangPt(P As tVec2, TangTo As tVec2)
    NP = NP + 1
    If NP = 1 Then
        FirstPoint = P
        FirstPointTang = TangTo
    End If
    NBia = NP - 1
    If NBia > UBound(BIARC) Then
        ReDim Preserve BIARC(NBia)
        Set BIARC(NBia) = New cBiARC
        If NP = 2 Then
            BIARC(NBia).SetPointsAndTangPts FirstPoint, FirstPointTang, _
                                            P, TangTo

        Else
            BIARC(NBia).SetPointsAndTangPts BIARC(NBia - 1).Point2, BIARC(NBia - 1).PtTangentTo2, _
                                            P, TangTo
        End If
    End If

    pvUpdatePathLength

End Sub
Friend Sub AddPointAndTangDirection(P As tVec2, TangDir As tVec2)
    AddPointAndTangPt P, SUM2(P, TangDir)
End Sub

Friend Sub SetPointTangStart(PointN As Long, TangentStart As tVec2, Optional IsDirection As Boolean = False)
    If PointN > NP Then Exit Sub
    If IsDirection Then
        BIARC(PointN).PtTangentTo1 = SUM2(BIARC(PointN).Point1, TangentStart)
    Else
        BIARC(PointN).PtTangentTo1 = TangentStart
    End If
    BIARC(PointN).CALC
    pvUpdatePathLength
End Sub
Friend Sub SetPointTangArrive(PointN As Long, TangentArrive As tVec2, Optional IsDirection As Boolean = False)
    If PointN < 2 Then Exit Sub                  '1st point hasnt arrive tangent
    If PointN > NP Then Exit Sub

    If IsDirection Then
        BIARC(PointN - 1).PtTangentTo2 = SUM2(BIARC(PointN - 1).Point2, TangentArrive)
    Else
        BIARC(PointN - 1).PtTangentTo2 = TangentArrive
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

Friend Sub Draw(CC As cCairoContext, Color As Long, Alpha As Double, Width As Double, _
                Optional Lines As Boolean = False, Optional FullCircles As Boolean = False)
    Dim I&
    For I = 1 To NBia
        BIARC(I).Draw CC, Color, Alpha, Width, Lines, FullCircles
    Next

End Sub

Friend Sub DrawOnlyCairoArcs(CC As cCairoContext)
    Dim I&
    'CC.MoveTo BIARC(1).Point1.x, BIARC(1).Point1.Y
    For I = 1 To NBia: BIARC(I).DrawArcsReadyToFIllStroke CC: Next
End Sub

Friend Sub DrawFilled(CC As cCairoContext, Color As Long, Alpha As Double)
    CC.SetSourceColor Color, Alpha
    DrawOnlyCairoArcs CC
    CC.Fill
End Sub



Friend Function InterpolatedPointAt(ByVal T As Double) As tVec2
    Dim currT     As Double
    Dim I&, J&
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
