VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "BiArc Interpolation"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cCollision 
      Caption         =   "Test Collision"
      Height          =   975
      Left            =   8040
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CheckBox chkInterpolate 
      Caption         =   "Interpolate"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   8040
      Top             =   1800
   End
   Begin VB.CommandButton cCF 
      Caption         =   "Test BiarcPATH"
      Height          =   975
      Left            =   8040
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
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
      Caption         =   "Click and Drag Control Points"
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BIARC     As cBiARC

Private T         As Double

Private SRF       As cCairoSurface
Private CC        As cCairoContext
Private CP        As cControlPoint
Private cPTS      As cControlPoints

Private PicHDC    As Long

Private Sub cCF_Click()
    chkInterpolate.value = vbUnchecked

    fCurveFit.Show vbModal
End Sub

Private Sub cCollision_Click()
    chkInterpolate.value = vbUnchecked

    fCollision.Show vbModal
End Sub

Private Sub chkInterpolate_Click()
    Timer1.Enabled = chkInterpolate = vbChecked

End Sub

Private Sub Form_Activate()

    BIARC.DRAW CC, vbYellow, 1, 3


    cPTS.DRAW CC
    SRF.DrawToDC PicHDC
    DoEvents
    SRF.DrawToDC PicHDC

End Sub

Private Sub Form_Load()
    Set BIARC = New cBiARC

    Set SRF = Cairo.CreateSurface(fMain.PIC.Width, fMain.PIC.Height, ImageSurface)    'size of our rendering-area in Pixels
    Set CC = SRF.CreateContext                   'create a Drawing-Context from the PixelSurface above

    CC.AntiAlias = CAIRO_ANTIALIAS_BEST
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.SetLineJoin CAIRO_LINE_JOIN_ROUND
    CC.SetLineWidth 1, True
    CC.SelectFont "Courier New", 9, vbGreen

    PicHDC = fMain.PIC.hDC

    Set cPTS = New cControlPoints

    BIARC.SetPointsAndControlPts Vec2(80, 120), Vec2(80 + 20, 120 + 40), _
                                 Vec2(300, 150), Vec2(300 + 20, 150 - 40)

    cPTS.Add "C1", BIARC.Point1.X, BIARC.Point1.Y, vbGreen, 14, 0.7
    cPTS.Add "C2", BIARC.Point2.X, BIARC.Point2.Y, vbRed, 14, 0.7
    cPTS.Add "T1", BIARC.ControlPt1.X, BIARC.ControlPt1.Y, vbGreen, 9, 0.7
    cPTS.Add "T2", BIARC.ControlPt2.X, BIARC.ControlPt2.Y, vbRed, 9, 0.7

    BIARC.CALC

    '    cCF_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim CP      As cControlPoint
    Set CP = cPTS.CheckControlPointUnderCursor(X, Y)
    If Not CP Is Nothing Then
        CP.SetMouseDownPoint X, Y:               ' BIARC.CalcAndDRAWBiARC P1, P2, NT1, NT2, 0    ': RENDERrc    ' RaiseEvent RefreshContents(CC)
        BIARC.CALC

    End If

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CP As cControlPoint, MOverStateChanged As Boolean
    If Button Then
        Set CP = cPTS.CheckControlPointUnderCursor(X, Y, True, MOverStateChanged)
        If Not CP Is Nothing Then

            Select Case CP.Key
            Case "C1"
                BIARC.Point1 = Vec2(CP.X, CP.Y)
            Case "C2"
                BIARC.Point2 = Vec2(CP.X, CP.Y)
            Case "T1"
                BIARC.ControlPt1 = Vec2(CP.X, CP.Y)
            Case "T2"
                BIARC.ControlPt2 = Vec2(CP.X, CP.Y)
            End Select



            BIARC.CALC
            With CC: .SetSourceColor 0: .Paint: End With
            BIARC.DRAW CC, vbYellow, 1, 3, True, True

            cPTS.DRAW CC
            SRF.DrawToDC PicHDC
            DoEvents
            '--------------------------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        End If
    End If

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cPTS.EnsureMouseUpState

End Sub

Private Sub Timer1_Timer()
    Dim P         As tVec2

    P = BIARC.InterpolatedPointAt(T)
    T = T + 0.02007
    If T > 1 Then T = T - 1
    BIARC.CALC
    With CC: .SetSourceColor 0: .Paint: End With
    BIARC.DRAW CC, vbYellow, 1, 3, True, True

    CC.Arc P.X, P.Y, 8
    CC.Stroke

    cPTS.DRAW CC
    SRF.DrawToDC PicHDC
    DoEvents
End Sub
