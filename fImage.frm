VERSION 5.00
Begin VB.Form fImage 
   Caption         =   "Image Biarc"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrollScale 
      Height          =   375
      Index           =   1
      Left            =   13200
      Max             =   200
      Min             =   25
      SmallChange     =   25
      TabIndex        =   5
      Top             =   5520
      Value           =   100
      Width           =   2055
   End
   Begin VB.HScrollBar scrollScale 
      Height          =   375
      Index           =   0
      Left            =   13200
      Max             =   200
      Min             =   25
      SmallChange     =   25
      TabIndex        =   4
      Top             =   4920
      Value           =   100
      Width           =   2055
   End
   Begin VB.ComboBox cTEST 
      Height          =   330
      Left            =   13200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chS 
      Caption         =   "Stretchable"
      Height          =   615
      Left            =   13200
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Timer TimerFPS 
      Interval        =   1000
      Left            =   14400
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   14400
      Top             =   2520
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   865
      TabIndex        =   0
      ToolTipText     =   "Change image"
      Top             =   120
      Width           =   12975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TEST"
      Height          =   255
      Left            =   13200
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "fImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BiArc         As cBiARC
Dim SRF2          As cCairoSurface
Dim CC2           As cCairoContext


Dim pW            As Long
Dim pH            As Long

Dim fps&

Dim TestMode      As Long

Dim BAPath        As cBiArcPath

Dim Stretchable   As Boolean



Private Sub chS_Click()
    Stretchable = chS = vbChecked
End Sub



Private Sub cTEST_Click()

    TestMode = cTEST.ListIndex
End Sub


Private Sub Form_Load()


    cTEST.AddItem "Simple BiArc"
    cTEST.AddItem "Biarc Path"
    cTEST.AddItem "Biarc Path AA(WIP)"
    cTEST.ListIndex = 0

    Set BiArc = New cBiARC


    pW = Me.PIC.Width - 1
    pH = Me.PIC.Height - 1
    Set SRF2 = Cairo.CreateSurface(pW + 1, pH + 1, ImageSurface)    'size of our rendering-area in Pixels


    Set CC2 = SRF2.CreateContext                 'create a Drawing-Context from the PixelSurface above
    CC2.AntiAlias = CAIRO_ANTIALIAS_BEST
    CC2.SetLineCap CAIRO_LINE_CAP_ROUND
    CC2.SetLineJoin CAIRO_LINE_JOIN_ROUND




    With BiArc
        .Point1 = Vec2(70, 70)
        .ControlPt1 = Vec2(200, 100)
        .Point2 = Vec2(460, 370)
        .ControlPt2 = Vec2(550, 377)
        .CALC
        .DRAW CC2, vbRed, 1, 4
        SRF2.DrawToDC PIC.hDC

        .SetImage App.Path & "\res\Rope.png", SRF2, 1, 0.75
    End With

    '------------------------------------------------------

    Set BAPath = New cBiArcPath

    With BAPath
        .SetupImage App.Path & "\res\Rope.png", SRF2, 0.5, 0.5
        .AddPointAndTangDirection Vec2(120, 170), Vec2(1, -1)
        .AddPointAndTangDirection Vec2(pW - 150, 170), Vec2(1, 1)
        .AddPointAndTangDirection Vec2(pW - 150, 440), Vec2(-1, 1)
        .AddPointAndTangDirection Vec2(120, 440), Vec2(-1, -1)

        .ClosePath


    End With





End Sub


Private Sub Form_Activate()
    PIC_Click
    Timer1.Enabled = True

End Sub
Private Sub scrollScale_Change(Index As Integer)
    BAPath.SetupImageScale fImage.scrollScale(0) * 0.01, fImage.scrollScale(1) * 0.01
End Sub

Private Sub scrollScale_Scroll(Index As Integer)
    BAPath.SetupImageScale fImage.scrollScale(0) * 0.01, fImage.scrollScale(1) * 0.01
End Sub
Private Sub PIC_Click()
    Static I&
    Randomize Timer

    Dim ScaleX#
    Dim ScaleY#

    ScaleX = fImage.scrollScale(0) * 0.01
    ScaleY = fImage.scrollScale(1) * 0.01


    If Timer1.Enabled Then
        I = (I + 1) Mod 8
        Select Case I
        Case 0
            BiArc.SetImage App.Path & "\res\Barbwire.png", SRF2
        Case 1
            BiArc.SetImage App.Path & "\res\Chain.png", SRF2
        Case 2
            BiArc.SetImage App.Path & "\res\Rope.png", SRF2
        Case 3
            BiArc.SetImage App.Path & "\res\road1.png", SRF2
        Case 4
            BiArc.SetImage App.Path & "\res\Smoke3.png", SRF2
        Case 5
            BiArc.SetImage App.Path & "\res\Telephonecord.png", SRF2
        Case 6
            BiArc.SetImage App.Path & "\res\Thorns.png", SRF2
        Case 7
            BiArc.SetImage App.Path & "\res\FloralStripe.png", SRF2

        End Select


        Select Case I
        Case 0
            BAPath.SetupImage App.Path & "\res\Barbwire.png", SRF2, ScaleX, ScaleY
        Case 1
            BAPath.SetupImage App.Path & "\res\Chain.png", SRF2, ScaleX, ScaleY
        Case 2
            BAPath.SetupImage App.Path & "\res\Rope.png", SRF2, ScaleX, ScaleY
        Case 3
            BAPath.SetupImage App.Path & "\res\road1.png", SRF2, ScaleX, ScaleY
        Case 4
            BAPath.SetupImage App.Path & "\res\Smoke3.png", SRF2, ScaleX, ScaleY
        Case 5
            BAPath.SetupImage App.Path & "\res\Telephonecord.png", SRF2, ScaleX, ScaleY
        Case 6
            BAPath.SetupImage App.Path & "\res\Thorns.png", SRF2, ScaleX, ScaleY
        Case 7
            BAPath.SetupImage App.Path & "\res\FloralStripe.png", SRF2, ScaleX, ScaleY

        End Select

        'FloralStripe

    End If



End Sub



Private Sub Timer1_Timer()
    Dim V         As tVec2
    If TestMode <> 0 Then

        If TestMode = 1 Then
            BAPath.DrawIMAGE True, IIf(Stretchable, 2000, 0)
        Else
            BAPath.DrawIMAGEAA True, IIf(Stretchable, 2000, 0)

        End If
        '        BAPath.DrawImageRC

        SRF2.DrawToDC PIC.hDC
        With BAPath

            '        V.X = V.X + (Rnd * 2 - 1) * 2
            '        V.Y = V.Y + (Rnd * 2 - 1) * 2
            V.X = pW * 0.5 + pW * 0.45 * Cos(Timer * 0.217)
            V.Y = pH * 0.5 + pH * 0.45 * Sin(Timer * 0.217)
            .SetPoint 2, V, True


            V.X = pW * 0.5 + pW * 0.33 * Cos(Timer * 0.231)
            V.Y = pH * 0.5 + pH * 0.33 * Sin(Timer * 0.231)
            .SetPoint 4, V, True



        End With

    Else

        With BiArc
            V = .Point1
            V.Y = pH * 0.5 + pH * 0.35 * Cos(Timer * 0.3)
            .Point1 = V

            V = .ControlPt1
            '        V.X = V.X + (Rnd * 2 - 1) * 2
            '        V.Y = V.Y + (Rnd * 2 - 1) * 5
            V.X = pW * 0.5 + pW * 0.5 * Cos(Timer * 0.27)
            V.Y = pH * 0.5 + pH * 0.5 * Sin(Timer * 0.27)
            .ControlPt1 = V

            V = .ControlPt2
            '        V.X = V.X + (Rnd * 2 - 1) * 2
            '        V.Y = V.Y + (Rnd * 2 - 1) * 2
            V.X = pW * 0.5 + pW * 0.5 * Cos(Timer * 0.21)
            V.Y = pH * 0.5 + pH * 0.5 * Sin(Timer * 0.21)
            .ControlPt2 = V

            V = .Point2
            '        V.X = V.X + (Rnd * 2 - 1) * 5
            '        V.Y = V.Y + (Rnd * 2 - 1) * 5

            V.X = pW * 0.75 + pW * 0.2 * Cos(Timer * 0.2)
            V.Y = pH * 0.5 + pW * 0.2 * Sin(Timer * 0.2)
            .Point2 = V

            .CALC

            '        'DrawIMAGE1 True
            '        DrawIMAGE2 True
            BiArc.DrawIMAGE CC2, True, IIf(Stretchable, 800, 0)
            SRF2.DrawToDC PIC.hDC
        End With

    End If




    fps = fps + 1

End Sub
'''
'''Private Sub DrawIMAGE1(ClearScreen As Boolean)
'''    Dim ArrSource() As Long
'''    Dim ArrSrf2() As Long
'''
'''    Dim W         As Long
'''    Dim H         As Long
'''    Dim X         As Double
'''    Dim Y         As Double
'''    Dim tx#, ty#
'''    Dim a#, R#, T As tVec2, wA&
'''    Dim cen       As tVec2
'''    Dim CosA#, SinA#
'''    Dim Segno#
'''    Dim currR#
'''    Dim StepX#
'''    Dim Repeat    As Double
'''
'''    Dim ScaleH    As Double
'''
'''
'''    If ClearScreen Then CC2.SetSourceColor 0: CC2.Paint
'''
'''
'''    iSRF.BindToArrayLong ArrSource
'''    SRF2.BindToArrayLong ArrSrf2
'''
'''    W = UBound(ArrSource, 1)
'''    H = UBound(ArrSource, 2)
'''
'''
'''    Repeat = 3.5
'''    ScaleH = 0.75
'''
'''
'''    With BiArc
'''
'''        'Repeat = .LengthTot / W 'AutoRepeat
'''
'''        StepX = 0.9 * (W * Repeat) / .LengthTot
'''
'''
'''        For X = 0 To W * Repeat Step StepX
'''
'''            T = .InterpolatedPointAt((X / W) / Repeat, wA)
'''            If wA = 1 Then
'''                cen = .Center1
'''                R = .Radius1
'''                Segno = IIf(.ClockWise1, 1, -1)
'''            Else
'''                cen = .Center2
'''                R = .Radius2
'''                Segno = IIf(.ClockWise2, 1, -1)
'''            End If
'''
'''            a = Atan2(T.X - cen.X, T.Y - cen.Y)
'''
'''            CosA = Cos(a)
'''            SinA = Sin(a)
'''
'''            For Y = 0 To H Step 0.5
'''                currR = (Segno * ScaleH * (Y - H * 0.5) + R)
'''
'''                tx = cen.X + CosA * currR
'''                ty = cen.Y + SinA * currR
'''                If tx >= 0 Then
'''                    If ty >= 0 Then
'''                        If tx <= pW Then
'''                            If ty <= pH Then
'''                                ArrSrf2(tx, ty) = ArrSource(X Mod W, Y)
'''                            End If
'''                        End If
'''                    End If
'''                End If
'''            Next
'''        Next
'''
'''        CC2.SetSourceColor vbCyan, 0.25
'''        CC2.Arc .ControlPt1.X, .ControlPt1.Y, 7
'''        CC2.MoveTo .ControlPt1.X, .ControlPt1.Y: CC2.LineTo .Point1.X, .Point1.Y
'''        CC2.Stroke
'''        CC2.Arc .ControlPt2.X, .ControlPt2.Y, 7
'''        CC2.MoveTo .ControlPt2.X, .ControlPt2.Y: CC2.LineTo .Point2.X, .Point2.Y
'''        CC2.Stroke
'''
'''    End With
'''
'''    iSRF.ReleaseArrayLong ArrSource
'''    SRF2.ReleaseArrayLong ArrSrf2
'''
'''
'''
'''    SRF2.DrawToDC PIC.hDC
'''
'''End Sub
'''
'''
'''
'''
'''Private Sub DrawIMAGE2(ClearScreen As Boolean)
'''    Dim ArrSource() As Long
'''    Dim ArrSrf2() As Long
'''    '    Dim ArrSource() As Byte
'''    '    Dim ArrSrf2() As Byte
'''
'''    Dim H2&, H22#
'''    Dim DX#, DY#
'''    Dim Sx#, Sy#
'''    Dim X&, Y&
'''    Dim xF&, xT&
'''    Dim yF&, yT&
'''    Dim D#
'''    Dim W&, H&
'''    Dim a#
'''    Dim GetX#, GetY#
'''    Dim bA As tVec2
'''    Dim bB As tVec2
'''
'''
'''    Dim dmin1#, dmax1#
'''    Dim dmin2#, dmax2#
'''
'''    If ClearScreen Then CC2.SetSourceRGB 0.2, 0.25, 0.25: CC2.Paint
'''
'''    iSRF.BindToArrayLong ArrSource
'''    SRF2.BindToArrayLong ArrSrf2
'''    '    iSRF.BindToArray ArrSource
'''    '    SRF2.BindToArray ArrSrf2
'''
'''    W = UBound(ArrSource, 1)    '\ 4
'''    H = UBound(ArrSource, 2)
'''
'''    H2 = H / 2
'''    H22 = H2 * H2
'''
'''    With BiArc
'''
'''        .GetBounds bA, bB
'''
'''        xF = bA.X - H2
'''        yF = bA.Y - H2
'''        xT = bB.X + H2
'''        yT = bB.Y + H2
'''
'''        If xF < 0 Then xF = 0
'''        If yF < 0 Then yF = 0
'''        If xT > pW Then xT = pW
'''        If yT > pH Then yT = pH
'''
'''        If xT < 0 Then xT = 0
'''        If yT < 0 Then yT = 0
'''        If xF > pW Then xF = pW
'''        If yF > pH Then yF = pH
'''
'''
'''
'''        dmin1 = .Radius1 - H2
'''        dmin1 = Sgn(dmin1) * dmin1 * dmin1
'''        dmax1 = .Radius1 + H2
'''        dmax1 = dmax1 * dmax1
'''
'''
'''        dmin2 = .Radius2 - H2
'''        dmin2 = Sgn(dmin2) * dmin2 * dmin2
'''        dmax2 = .Radius2 + H2
'''        dmax2 = dmax2 * dmax2
'''
'''        For X = xF To xT
'''            For Y = yF To yT
'''                DX = X - .Center1.X
'''                DY = Y - .Center1.Y
'''                'D = Sqr(DX * DX + DY * DY) - .Radius1
'''                'If Abs(D) < H2 Then
'''                D = DX * DX + DY * DY
'''                If D > dmin1 And D < dmax1 Then
'''                    D = Sqr(D) - .Radius1
'''
'''                    a = Atan2(DX, DY)
'''                    If IsAngBetween(a, .A11, .A12) Then
'''                        GetY = H2 - D * IIf(.ClockWise1, 1, -1)
'''                        If .ClockWise1 Then
'''                            GetX = AngleDIFF(a, .A11) * .Radius1
'''                        Else
'''                            GetX = -AngleDIFF(a, .A12) * .Radius1
'''                        End If
'''                        GetX = GetX Mod W
'''
'''                        ArrSrf2(X, Y) = ArrSource(GetX, GetY Mod H)
'''                        'ArrSrf2(X * 4, Y) = ArrSource(GetX * 4, GetY Mod H)
'''
'''                    End If
'''                End If
'''                DX = X - .Center2.X
'''                DY = Y - .Center2.Y
'''                'D = Sqr(DX * DX + DY * DY) - .Radius2
'''                'If Abs(D) < H2 Then
'''                D = DX * DX + DY * DY
'''                If D > dmin2 And D < dmax2 Then
'''                    D = Sqr(D) - .Radius2
'''                    a = Atan2(DX, DY)
'''                    If IsAngBetween(a, .A21, .A22) Then
'''                        GetY = H2 - D * IIf(.ClockWise2, 1, -1)
'''
'''                        If .ClockWise2 Then
'''                            GetX = AngleDIFF(a, .A21) * .Radius2
'''                        Else
'''                            GetX = -AngleDIFF(a, .A22) * .Radius2
'''                        End If
'''                        GetX = GetX + .LengthArc1
'''                        GetX = GetX Mod W
'''                        ArrSrf2(X, Y) = ArrSource(GetX, GetY Mod H)
'''                        'ArrSrf2(X * 4, Y) = ArrSource(GetX * 4, GetY Mod H)
'''                    End If
'''                End If
'''
'''            Next
'''        Next
'''
'''
'''    End With
'''
'''    iSRF.ReleaseArrayLong ArrSource
'''    SRF2.ReleaseArrayLong ArrSrf2
'''
'''    '    iSRF.ReleaseArray ArrSource
'''    '    SRF2.ReleaseArray ArrSrf2
'''
'''
'''    SRF2.DrawToDC PIC.hDC
'''
'''
'''End Sub

Private Sub TimerFPS_Timer()

    fImage.Caption = "FPS " & fps


    fps = 0

End Sub
