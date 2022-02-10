VERSION 5.00
Begin VB.Form fCurveFit 
   Caption         =   "Path test"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShapes 
      Caption         =   "Shapes"
      Height          =   975
      Left            =   8040
      TabIndex        =   2
      ToolTipText     =   "Click to see test Shapes"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   8280
      Top             =   1920
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   441
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Click Image to Randomize"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "fCurveFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BI()          As cBiARC
Dim NB            As Long
Dim P()           As tVec2
Dim NP            As Long

Dim SRF2          As cCairoSurface
Dim CC2           As cCairoContext

Dim BiarcPath     As cBiArcPath

Dim T             As Double

Dim TS            As Long


Private Sub Form_Activate()
    TEST2
End Sub

Public Sub TEST()


    Dim I&, J&
    Dim C1        As tVec2
    Dim C2        As tVec2
    Dim R1#, R2#
    Dim A11#, A12#, A21#, A22#

    NP = 6                                       ' min 4 - and Multpile of 2

    Dim P1        As tVec2
    Dim P2        As tVec2
    Dim T1        As tVec2
    Dim T2        As tVec2

    ReDim P(NP)
    For I = 1 To NP
        P(I).X = 50 + Rnd * 300
        P(I).Y = 50 + Rnd * 300
    Next

    '4-1
    '6-2
    '8-3
    '10-4
    '12-5

    NB = NP \ 2 - 1

    fCurveFit.Caption = "Np = " & NP & "   NB = " & NB

    ReDim BI(NB)
    With CC2: .SetSourceColor 0: .Paint: End With
    J = 1
    For I = 1 To NB
        Set BI(I) = New cBiARC
        '        P1 = P(J): J = J + 1
        '        If I = 1 Then
        '            T1 = P(J): J = J + 2
        '        Else
        '            T1 = SUM2(P(J - 1), Vec2(BI(I - 1).TangentDir2.x, BI(I - 1).TangentDir2.Y))
        '            J = J + 2
        '        End If
        '        P2 = P(J): J = J - 1
        '        T2 = SUM2(P2, DIFF2(P2, P(J)))

        P1 = P(J): J = J + 1
        T1 = P(J): J = J + 1
        P2 = P(J): J = J + 1
        T2 = P(J)

        BI(I).SetPointsAndControlPts P1, T1, P2, T2
        BI(I).CALC
        BI(I).DRAW CC2, vbYellow, 1, 3, True

        '        J = J + 1
        J = J - 1

    Next


    For I = 1 To NP
        If I Mod 2 = 0 Then CC2.SetSourceColor vbRed Else: CC2.SetSourceColor vbYellow
        CC2.Arc P(I).X, P(I).Y, 8
        CC2.Fill
        CC2.TextOut P(I).X - 4, P(I).Y - 7, CStr(I)
    Next


    SRF2.DrawToDC PIC.hDC

End Sub

Private Sub Form_Load()
    Randomize Timer

    Set SRF2 = Cairo.CreateSurface(Me.PIC.Width, Me.PIC.Height, ImageSurface)    'size of our rendering-area in Pixels
    Set CC2 = SRF2.CreateContext                 'create a Drawing-Context from the PixelSurface above
    CC2.AntiAlias = CAIRO_ANTIALIAS_BEST
    CC2.SetLineCap CAIRO_LINE_CAP_ROUND
    CC2.SetLineJoin CAIRO_LINE_JOIN_ROUND
End Sub

Private Sub PIC_Click()
'TEST
    TEST2
End Sub


Private Sub TEST2()
    Me.Caption = "BiArc-PATH Test"

    Dim I&
    With CC2: .SetSourceColor 0: .Paint: End With

    Set BiarcPath = New cBiArcPath

    With BiarcPath
        '        .AddPointAndControlPoint Vec2(50, 50), Vec2(70, 70)
        '        .AddPointAndControlPoint Vec2(85, 80), Vec2(115, 85)
        '        .AddPointAndControlPoint Vec2(290, 180), Vec2(150, 185)
        '        .SetPointTangArrive 2, Vec2(0, 1), True
        '        .SetPointTangArrive 3, Vec2(10, 0), True
        '        .SetPointTangStart 2, Vec2(2, 0), True
        '        .AddPointAndTangDirection Vec2(100, 350), Vec2(0, 1)
        '        .AddPointAndTangDirection Vec2(300, 350), Vec2(1, 0)
        '        .DRAW CC2, vbYellow, 1, 3, True

        For I = 1 To 3 + Rnd * 1
            .AddPointAndTangDirection Vec2(100 + Rnd * 300, 100 + Rnd * 300), Vec2(Rnd * 2 - 1, Rnd * 2 - 1)
        Next

        If Rnd < 0.5 Then                        'Close Path
            .ClosePath
        End If

    End With

    SRF2.DrawToDC PIC.hDC

    Timer1.Enabled = True

    Me.Caption = "BiArc-PATH Test     N Points: " & BiarcPath.Npoints
End Sub

Private Sub Timer1_Timer()
    Dim P         As tVec2
    Dim I         As Long

    T = T + 0.01251
    If T > 1 Then T = T - 1

    With CC2: .SetSourceColor 0: .Paint: End With
    BiarcPath.DRAW CC2, vbYellow, 1, 3           ', True
    If BiarcPath.Closed Then
        CC2.SetSourceColor vbCyan, 0.33
        BiarcPath.DrawOnlyCairoArcs CC2
        CC2.Fill
    End If

    CC2.SetSourceColor vbGreen, 0.5
    For I = 1 To BiarcPath.Npoints
        P = BiarcPath.Point(I)
        CC2.Arc P.X, P.Y, 8
        CC2.Fill
        CC2.TextOut P.X - 4, P.Y - 7, CStr(I)
    Next

    CC2.SetSourceColor vbWhite
    P = BiarcPath.InterpolatedPointAt(T)
    CC2.Arc P.X, P.Y, 7.5
    CC2.Fill

    SRF2.DrawToDC PIC.hDC
    DoEvents
End Sub



Private Sub cmdShapes_Click()
    Dim X#, Y#
    Dim oX#, oY#
    Dim C         As Long
    Dim I         As Long
    Dim NewTang   As tVec2
    Dim Delta1    As tVec2
    Dim Delta2    As tVec2
    Dim A#, R#



    TS = 2                                       '<<<<<<<<<<<< force TS to 2 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    BiarcPath.Destroy

    Select Case TS
    Case 0
        With BiarcPath
            .AddPointAndTangDirection Vec2(250, 50), Vec2(0, 1)
            .AddPointAndTangDirection Vec2(160, 225), Vec2(0, 1)
            .AddPointAndTangDirection Vec2(250, 400), Vec2(0, 1)
            .AddPointAndTangDirection Vec2(340, 225), Vec2(0, -1)
            .ClosePath
            .SetPointTangStart 3, Vec2(0, -1), True
            .SetPointTangArrive 5, Vec2(0, -1), True
        End With
    Case 1
        With BiarcPath
            .AddPointAndTangDirection Vec2(250, 50), Vec2(-1, 0)
            .AddPointAndTangDirection Vec2(150, 200), Vec2(1, 1)
            .AddPointAndTangDirection Vec2(250, 420), Vec2(0, 1)
            .AddPointAndTangDirection Vec2(350, 200), Vec2(1, -1)
            .ClosePath
            .SetPointTangStart 3, Vec2(0, -1), True
        End With
    Case 2                                       ' Serie of Points Interpolation
        With BiarcPath
            C = 1
            '            For X = 5 To 500 Step 30
            '               Y = 250 + Cos(X * 0.18) * 80 - Rnd * 150
            '                If C Mod 2 = 0 Then
            '                    .AddPointAndControlPoint Vec2(oX, oY), Vec2(X, Y)
            '                End If
            '                oX = X
            '                oY = Y

            For A = 0 To PI2 Step 0.4
                R = (150 + (Rnd * 2 - 1) * 70)
                X = 250 + Cos(A) * R
                Y = 220 + Sin(A) * R
                .AddPointAndTangDirection Vec2(X, Y), Vec2(1, 0)
                C = C + 1
            Next

            .ClosePath
            .AutoSetTangents

        End With
    End Select
    TS = (TS + 1) Mod 3


End Sub

