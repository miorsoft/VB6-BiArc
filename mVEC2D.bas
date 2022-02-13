Attribute VB_Name = "mVEC2D"
Option Explicit
'*************************************************************************
'************************* V E C T O R S & MATHS  ************************
'*************************************************************************
Public Type tVec2
    X             As Double
    Y             As Double
End Type

Public Type tMAT2
    m00           As Double
    m01           As Double
    m10           As Double
    m11           As Double
End Type


Public Const PI   As Double = 3.14159265358979
Public Const PI2  As Double = 6.28318530717959
Public Const PIh  As Double = 1.5707963267949

Public Const EPSILON As Double = 0.00001
Public Const EPSILON_SQ As Double = EPSILON * EPSILON

Public Const MAX_VALUE As Double = 1E+32

Public Function Vec2(X As Double, Y As Double) As tVec2
    Vec2.X = X
    Vec2.Y = Y
End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.X = -V.X
    Vec2Negative.Y = -V.Y
End Function



Public Function Vec2ADD(V1 As tVec2, V2 As tVec2) As tVec2
    Vec2ADD.X = V1.X + V2.X
    Vec2ADD.Y = V1.Y + V2.Y
End Function
Public Function SUM2(V1 As tVec2, V2 As tVec2) As tVec2
    SUM2.X = V1.X + V2.X
    SUM2.Y = V1.Y + V2.Y
End Function
Public Function DIFF2(V1 As tVec2, V2 As tVec2) As tVec2
    DIFF2.X = V1.X - V2.X
    DIFF2.Y = V1.Y - V2.Y
End Function

Public Function Vec2MULV(V1 As tVec2, V2 As tVec2) As tVec2
    Vec2MULV.X = V1.X * V2.X
    Vec2MULV.Y = V1.Y * V2.Y
End Function
Public Function MUL2(V As tVec2, S As Double) As tVec2
    MUL2.X = V.X * S
    MUL2.Y = V.Y * S
End Function

Public Function ADDScaled2(V1 As tVec2, V2 As tVec2, S As Double) As tVec2
    ADDScaled2.X = V1.X + V2.X * S
    ADDScaled2.Y = V1.Y + V2.Y * S
End Function

Public Function LengthSq2(V As tVec2) As Double
    LengthSq2 = V.X * V.X + V.Y * V.Y
End Function

Public Function Length2(V As tVec2) As Double
    Length2 = Sqr(V.X * V.X + V.Y * V.Y)

End Function
Public Function Vec2Magnitude(V As tVec2) As Double
    Vec2Magnitude = Sqr(V.X * V.X + V.Y * V.Y)
End Function

Public Function SIDE(P As tVec2, L1 As tVec2, L2 As tVec2) As Double
'https://stackoverflow.com/questions/1560492/how-to-tell-whether-a-point-is-to-the-right-or-left-side-of-a-line
    SIDE = Sgn((L2.X - L1.X) * (P.Y - L1.Y) - (L2.Y - L1.Y) * (P.X - L1.X))
End Function

Public Function Vec2Rotate(V As tVec2, radians As Double) As tVec2
'real c = std::cos( radians );
'real s = std::sin( radians );

'real xp = x * c - y * s;
'real yp = x * s + y * c;

    Dim S         As Double
    Dim C         As Double
    C = Cos(radians)
    S = Sin(radians)

    Vec2Rotate.X = V.X * C - V.Y * S
    Vec2Rotate.Y = V.X * S + V.Y * C
End Function

Public Function Normalize2(V As tVec2) As tVec2
    Dim D         As Double
    D = Length2(V)
    If D Then
        D = 1# / D
        Normalize2.X = V.X * D
        Normalize2.Y = V.Y * D
    End If

End Function

'//******************************************************************************
'// Check if the vector length is within epsilon of 1
'//******************************************************************************
'bool Vec_IsNormalized_Eps(const tVec3 & value, float epsilon)
'{
'    const float sqrMag = Vec_DotProduct(value,value);
'    return      sqrMag >= (1.0f - epsilon)*(1.0f - epsilon)
'            &&  sqrMag <= (1.0f + epsilon)*(1.0f + epsilon);
'}

Public Function Vec_IsNormalized_Eps(value As tVec2) As Boolean

    Dim sqrMag    As Double
    sqrMag = DOT2(value, value)


    Vec_IsNormalized_Eps = (sqrMag >= (1# - EPSILON) * (1# - EPSILON)) _
                           And (sqrMag <= (1# + EPSILON) * (1# + EPSILON))

End Function



Public Function Vec2MIN(A As tVec2, B As tVec2) As tVec2
    Vec2MIN.X = IIf(A.X < B.X, A.X, B.X)
    Vec2MIN.Y = IIf(A.Y < B.Y, A.Y, B.Y)
End Function

Public Function Vec2MAX(A As tVec2, B As tVec2) As tVec2
    Vec2MAX.X = IIf(A.X > B.X, A.X, B.X)
    Vec2MAX.Y = IIf(A.Y > B.Y, A.Y, B.Y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function DOT2(A As tVec2, B As tVec2) As Double
    DOT2 = A.X * B.X + A.Y * B.Y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, A As Double) As tVec2
    Vec2CROSSva.X = A * V.Y
    Vec2CROSSva.Y = -A * V.X
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(A As Double, V As tVec2) As tVec2
    Vec2CROSSav.X = -A * V.Y
    Vec2CROSSav.Y = A * V.X
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(A As tVec2, B As tVec2) As Double
    Vec2CROSS = A.X * B.Y - A.Y * B.X
End Function

Public Function Vec2CROSS2(A As tVec2, B As tVec2) As tVec2
'    float x = lhs.m_y*rhs.m_z - lhs.m_z*rhs.m_y;
'    float y = lhs.m_z*rhs.m_x - lhs.m_x*rhs.m_z;
'    float z = lhs.m_x*rhs.m_y - lhs.m_y*rhs.m_x;


'    Vec2CROSS2.X = A.Y * B.X - A.X * B.Y
'    Vec2CROSS2.Y = A.X * B.Y - A.Y * B.X


    Vec2CROSS2.X = -(B.Y - A.Y)
    Vec2CROSS2.Y = (B.X - A.X)


    ''''    Vec2CROSS2.X = A.Y * B.z - A.z * B.Y
    ''''    Vec2CROSS2.Y = A.z * B.X - A.X * B.z
    ''''    Vec2CROSS2.z = A.X * B.Y - A.Y * B.X
    '''
    '''    Vec2CROSS2.X = A.Y * 1 - 1 * B.Y
    '''    Vec2CROSS2.Y = 1 * B.X - A.X * 1
    '''    'Vec2CROSS2.z = A.X * B.Y - A.Y * B.X


End Function


Public Function Vec2DISTANCEsq(A As tVec2, B As tVec2) As Double
    Dim Dx        As Double
    Dim Dy        As Double
    Dx = A.X - B.X
    Dy = A.Y - B.Y
    Vec2DISTANCEsq = Dx * Dx + Dy * Dy
End Function


'************************************************************************************



Public Function matTranspose(M As tMAT2) As tMAT2
    With M
        matTranspose.m00 = .m00
        matTranspose.m01 = .m10                  '
        matTranspose.m10 = .m01                  '
        matTranspose.m11 = .m11
    End With
End Function

Public Function matMULv(M As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );
    With M
        matMULv.X = .m00 * V.X + .m01 * V.Y
        matMULv.Y = .m10 * V.X + .m11 * V.Y
    End With

End Function

Public Function SetOrient(radians As Double) As tMAT2
'    real c = std::cos( radians );
'    real s = std::sin( radians );
'
'    m00 = c; m01 = -s;
'    m10 = s; m11 =  c;

    Dim C         As Double
    Dim S         As Double

    C = Cos(radians)
    S = Sin(radians)

    With SetOrient
        .m00 = C
        .m01 = -S
        .m10 = S
        .m11 = C
    End With

End Function


Public Function VectorProject(ByRef V As tVec2, ByRef Vto As tVec2) As tVec2
'Poject Vector V to vector Vto
    Dim K         As Double
    Dim D         As Double

    D = Vto.X * Vto.X + Vto.Y * Vto.Y
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    K = (V.X * Vto.X + V.Y * Vto.Y) * D

    VectorProject.X = (Vto.X * D) * K
    VectorProject.Y = (Vto.Y * D) * K

End Function
Public Function VectorProjectN(ByRef V As tVec2, ByRef VtoN As tVec2) As tVec2
'Poject Vector V to vector VtoN
    Dim K         As Double

    K = (V.X * VtoN.X + V.Y * VtoN.Y)

    VectorProjectN.X = VtoN.X * K
    VectorProjectN.Y = VtoN.Y * K

End Function

Public Function VectorReflect(ByRef V As tVec2, ByRef wall As tVec2) As tVec2
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot      As Double
    Dim D         As Double
    Dim NwX       As Double
    Dim NwY       As Double

    D = (wall.X * wall.X + wall.Y * wall.Y)
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    NwX = wall.X * D
    NwY = wall.Y * D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.X * NwX + V.Y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.X = -V.X + NwX
    VectorReflect.Y = -V.Y + NwY


End Function


Public Function ACos(X As Double) As Double
'    ACOS = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    ACos = Atn(-X / Sqr(-X * X + 1#)) + 2# * PIh
End Function
Public Function ASin(ByVal X As Double) As Double
    ASin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function AngleDIFF(ByVal A1 As Double, ByVal A2 As Double) As Double

    AngleDIFF = A1 - A2
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend

End Function

Public Function IsEqualEps(ByVal A As Double, ByVal B As Double) As Boolean
    IsEqualEps = (Abs(A - B) < EPSILON)
End Function

Public Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
    If X Then                                    '''Sempre USATA
        Atan2 = -PI + Atn(Y / X) - (X > 0!) * PI
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function

''
' Divides two integers, placing the remainder in a supplied variable.
'
' @param a The dividend.
' @param b The divosor.
' @param Remainder The variable to place the remainder of the division.
' @return The quotient of the division.
'
Public Function DivRem(ByVal A As Long, ByVal B As Long, ByRef remainder As Long) As Long
    DivRem = A \ B
    remainder = A - (B * DivRem)                 ' this is about 2x faster than Mod.
End Function

Public Function LogBase(ByVal D As Double, ByVal NewBase As Double) As Double
    LogBase = Log(D) / Log(NewBase)
End Function

Public Function ToString2(V As tVec2) As String
    ToString2 = Format(V.X, "#.###") & " , " & Format(V.Y, "#.###")
End Function
Public Function IsAngBetween(ByVal MiddleA#, ByVal StartA#, ByVal EndA#) As Boolean
    Dim T#
    'https://math.stackexchange.com/questions/1044905/simple-angle-between-two-angles-of-circle
    If StartA < 0 Then StartA = StartA + PI2
    If EndA < 0 Then EndA = EndA + PI2
    If MiddleA < 0 Then MiddleA = MiddleA + PI2

    MiddleA = MiddleA - StartA
    EndA = EndA - StartA
    If MiddleA < 0 Then MiddleA = MiddleA + PI2
    If EndA < 0 Then EndA = EndA + PI2
    IsAngBetween = MiddleA < EndA

End Function


'
'
'/*
'   Linear Regression
'   y(x) = a + b x, for n samples
'   The following assumes the standard deviations are unknown for x and y
'   Return a, b and r the regression coefficient
'*/
'int LinRegress(double *x,double *y,int n,double *a,double *b,double *r)
'{
'   int i;
'   double sumx=0,sumy=0,sumx2=0,sumy2=0,sumxy=0;
'   double sxx,syy,sxy;
'
'   *a = 0;
'   *b = 0;
'   *r = 0;
'   if (n < 2)
'      return(FALSE);
'
'   /* Conpute some things we need */
'   for (i=0;i<n;i++) {
'      sumx += x[i];
'      sumy += y[i];
'      sumx2 += (x[i] * x[i]);
'      sumy2 += (y[i] * y[i]);
'      sumxy += (x[i] * y[i]);
'   }
'   sxx = sumx2 - sumx * sumx / n;
'   syy = sumy2 - sumy * sumy / n;
'   sxy = sumxy - sumx * sumy / n;
'
'   /* Infinite slope (b), non existant intercept (a) */
'   if (ABS(sxx) == 0)
'      return(FALSE);
'
'   /* Calculate the slope (b) and intercept (a) */
'   *b = sxy / sxx;
'   *a = sumy / n - (*b) * sumx / n;
'
'   /* Compute the regression coefficient */
'   if (ABS(syy) == 0)
'      *r = 1;
'   Else
'      *r = sxy / sqrt(sxx * syy);
'
'   return(TRUE);
'}
'



'   Linear Regression
'   y(x) = a + b x, for n samples
'   The following assumes the standard deviations are unknown for x and y
'   Return a, b and r the regression coefficient
Public Sub LinRegress(X() As Double, Y() As Double, ra#, rb#, rr#)
    Dim I&, N&
    Dim sumX#, sumY#
    Dim sumX2#, sumY2#, sumXY#
    Dim sXX#, sYY#, sXY#

    N = UBound(X)
    ra = 0: rb = 0: rr = 0
    '/* Conpute some things we need */
    For I = 1 To N
        sumX = sumX + X(I)
        sumY = sumY + Y(I)
        sumX2 = sumX2 + (X(I) * X(I))
        sumY2 = sumY2 + (Y(I) * Y(I))
        sumXY = sumXY + (X(I) * Y(I))
    Next
    sXX = sumX2 - sumX * sumX / N
    sYY = sumY2 - sumY * sumY / N
    sXY = sumXY - sumX * sumY / N

    '   /* Infinite slope (b), non existant intercept (a) */
    If (Abs(sXX) = 0) Then
        Stop
        '      return(FALSE);
    End If
    '/* Calculate the slope (b) and intercept (a) */
    rb = sXY / sXX
    ra = sumY / N - (rb) * sumX / N

    '/* Compute the regression coefficient */
    If (Abs(sYY) = 0) Then
        rr = 1
    Else
        rr = sXY / Sqr(sXX * sYY)
    End If

End Sub

'   Linear Regression
'   y(x) = a + b x, for n samples
'   The following assumes the standard deviations are unknown for x and y
'   Return a, b and r the regression coefficient
Public Function LinRegress3(P1 As tVec2, P2 As tVec2, P3 As tVec2, ra#, rb#, rr#) As tVec2
    Dim I&, N&, InvN#
    Dim sumX#, sumY#
    Dim sumX2#, sumY2#, sumXY#
    Dim sXX#, sYY#, sXY#

    N = 3: InvN = 1 / N
    ra = 0: rb = 0: rr = 0
    '/* Conpute some things we need */
    sumX = P1.X + P2.X + P3.X
    sumY = P1.Y + P2.Y + P3.Y
    sumX2 = P1.X * P1.X + P2.X * P2.X + P3.X * P3.X
    sumY2 = P1.Y * P1.Y + P2.Y * P2.Y + P3.Y * P3.Y
    sumXY = P1.X * P1.Y + P2.X * P2.Y + P3.X * P3.Y
    sXX = sumX2 - sumX * sumX * InvN
    sYY = sumY2 - sumY * sumY * InvN
    sXY = sumXY - sumX * sumY * InvN

    '   /* Infinite slope (b), non existant intercept (a) */
    If (Abs(sXX) = 0) Then
        Stop
        '      return(FALSE);
        LinRegress3 = Vec2(0, 1)
    End If
    '/* Calculate the slope (b) and intercept (a) */
    rb = sXY / sXX
    ra = sumY * InvN - (rb) * sumX * InvN

    LinRegress3 = Normalize2(Vec2(sXX, sXY))
    '/* Compute the regression coefficient */
    If (Abs(sYY) = 0) Then
        rr = 1
    Else
        rr = sXY / Sqr(sXX * sYY)
    End If

End Function

Public Function Rotate90(V As tVec2) As tVec2
    With V
        Rotate90.X = -V.Y
        Rotate90.Y = V.X
    End With
End Function

