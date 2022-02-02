Attribute VB_Name = "mVEC2D"
Option Explicit
'*************************************************************************
'************************* V E C T O R S & MATHS  ************************
'*************************************************************************
Public Type tVec2
    x             As Double
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

Public Const EPSILON As Double = 0.001           '0.0001
Public Const EPSILON_SQ As Double = EPSILON * EPSILON

Public Const MAX_VALUE As Double = 1E+32

Public Function Vec2(x As Double, Y As Double) As tVec2
    Vec2.x = x
    Vec2.Y = Y
End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.x = -V.x
    Vec2Negative.Y = -V.Y
End Function



Public Function Vec2ADD(V1 As tVec2, V2 As tVec2) As tVec2
    Vec2ADD.x = V1.x + V2.x
    Vec2ADD.Y = V1.Y + V2.Y
End Function
Public Function SUM2(V1 As tVec2, V2 As tVec2) As tVec2
    SUM2.x = V1.x + V2.x
    SUM2.Y = V1.Y + V2.Y
End Function
Public Function SUB2(V1 As tVec2, V2 As tVec2) As tVec2
    SUB2.x = V1.x - V2.x
    SUB2.Y = V1.Y - V2.Y
End Function

Public Function Vec2MULV(V1 As tVec2, V2 As tVec2) As tVec2
    Vec2MULV.x = V1.x * V2.x
    Vec2MULV.Y = V1.Y * V2.Y
End Function
Public Function MUL2(V As tVec2, S As Double) As tVec2
    MUL2.x = V.x * S
    MUL2.Y = V.Y * S
End Function

Public Function ADDScaled2(V1 As tVec2, V2 As tVec2, S As Double) As tVec2
    ADDScaled2.x = V1.x + V2.x * S
    ADDScaled2.Y = V1.Y + V2.Y * S
End Function

Public Function LengthSq2(V As tVec2) As Double
    LengthSq2 = V.x * V.x + V.Y * V.Y
End Function

Public Function Length2(V As tVec2) As Double
'   Length2 = FASTsqr(V.X * V.X + V.Y * V.Y)
    Length2 = Sqr(V.x * V.x + V.Y * V.Y)

End Function
Public Function Vec2Magnitude(V As tVec2) As Double
'   Length2 = FASTsqr(V.X * V.X + V.Y * V.Y)
    Vec2Magnitude = Sqr(V.x * V.x + V.Y * V.Y)

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

    Vec2Rotate.x = V.x * C - V.Y * S
    Vec2Rotate.Y = V.x * S + V.Y * C
End Function

Public Function Normalize2(V As tVec2) As tVec2
    Dim d         As Double
    d = Length2(V)
    If d Then
        d = 1# / d
        Normalize2.x = V.x * d
        Normalize2.Y = V.Y * d
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



Public Function Vec2MIN(a As tVec2, b As tVec2) As tVec2
    Vec2MIN.x = IIf(a.x < b.x, a.x, b.x)
    Vec2MIN.Y = IIf(a.Y < b.Y, a.Y, b.Y)
End Function

Public Function Vec2MAX(a As tVec2, b As tVec2) As tVec2
    Vec2MAX.x = IIf(a.x > b.x, a.x, b.x)
    Vec2MAX.Y = IIf(a.Y > b.Y, a.Y, b.Y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function DOT2(a As tVec2, b As tVec2) As Double
    DOT2 = a.x * b.x + a.Y * b.Y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, a As Double) As tVec2
    Vec2CROSSva.x = a * V.Y
    Vec2CROSSva.Y = -a * V.x
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(a As Double, V As tVec2) As tVec2
    Vec2CROSSav.x = -a * V.Y
    Vec2CROSSav.Y = a * V.x
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(a As tVec2, b As tVec2) As Double
    Vec2CROSS = a.x * b.Y - a.Y * b.x
End Function

Public Function Vec2CROSS2(a As tVec2, b As tVec2) As tVec2
'    float x = lhs.m_y*rhs.m_z - lhs.m_z*rhs.m_y;
'    float y = lhs.m_z*rhs.m_x - lhs.m_x*rhs.m_z;
'    float z = lhs.m_x*rhs.m_y - lhs.m_y*rhs.m_x;


'    Vec2CROSS2.X = A.Y * B.X - A.X * B.Y
'    Vec2CROSS2.Y = A.X * B.Y - A.Y * B.X


    Vec2CROSS2.x = -(b.Y - a.Y)
    Vec2CROSS2.Y = (b.x - a.x)


    ''''    Vec2CROSS2.X = A.Y * B.z - A.z * B.Y
    ''''    Vec2CROSS2.Y = A.z * B.X - A.X * B.z
    ''''    Vec2CROSS2.z = A.X * B.Y - A.Y * B.X
    '''
    '''    Vec2CROSS2.X = A.Y * 1 - 1 * B.Y
    '''    Vec2CROSS2.Y = 1 * B.X - A.X * 1
    '''    'Vec2CROSS2.z = A.X * B.Y - A.Y * B.X


End Function


Public Function Vec2DISTANCEsq(a As tVec2, b As tVec2) As Double
    Dim Dx        As Double
    Dim DY        As Double
    Dx = a.x - b.x
    DY = a.Y - b.Y
    Vec2DISTANCEsq = Dx * Dx + DY * DY
End Function


'************************************************************************************



Public Function matTranspose(m As tMAT2) As tMAT2
    With m
        matTranspose.m00 = .m00
        matTranspose.m01 = .m10                  '
        matTranspose.m10 = .m01                  '
        matTranspose.m11 = .m11
    End With
End Function

Public Function matMULv(m As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );
    With m
        matMULv.x = .m00 * V.x + .m01 * V.Y
        matMULv.Y = .m10 * V.x + .m11 * V.Y
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
    Dim d         As Double



    d = Vto.x * Vto.x + Vto.Y * Vto.Y
    If d = 0 Then Exit Function

    d = 1 / Sqr(d)

    K = (V.x * Vto.x + V.Y * Vto.Y) * d

    VectorProject.x = (Vto.x * d) * K
    VectorProject.Y = (Vto.Y * d) * K

End Function

Public Function VectorReflect(ByRef V As tVec2, ByRef wall As tVec2) As tVec2
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot      As Double
    Dim d         As Double
    Dim NwX       As Double
    Dim NwY       As Double

    d = (wall.x * wall.x + wall.Y * wall.Y)
    If d = 0 Then Exit Function

    d = 1 / Sqr(d)

    NwX = wall.x * d
    NwY = wall.Y * d
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.x * NwX + V.Y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.x = -V.x + NwX
    VectorReflect.Y = -V.Y + NwY


End Function


Public Function ACos(x As Double) As Double
'    ACOS = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    ACos = Atn(-x / Sqr(-x * x + 1#)) + 2# * PIh
End Function
Public Function ASin(ByVal x As Double) As Double
    ASin = Atn(x / Sqr(-x * x + 1))
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

Public Function IsEqualEps(ByVal a As Double, ByVal b As Double) As Boolean
    IsEqualEps = (Abs(a - b) < EPSILON)

End Function

Public Function Atan2(ByVal x As Double, ByVal Y As Double) As Double
    If x Then                                    '''Sempre USATA
        Atan2 = -PI + Atn(Y / x) - (x > 0!) * PI
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function

Public Function Atan22(ByVal x As Double, ByVal Y As Double) As Double
'Faster By Tanner   ''''Max error is , <= 0.005 radians
    If x Then
        Dim z     As Double
        z = Y / x
        If (Abs(z) < 1#) Then
            Atan22 = z / (1# + 0.28 * z * z)
            If (x < 0#) Then
                If (Y < 0#) Then
                    Atan22 = Atan22 - PI
                Else
                    Atan22 = Atan22 + PI
                End If
            End If
        Else
            Atan22 = PIh - z / (z * z + 0.28)
            If (Y < 0#) Then Atan22 = Atan22 - PI
        End If
    Else
        If (Y > 0#) Then
            Atan22 = PIh
        ElseIf (Y = 0#) Then
            Atan22 = 0#
        Else
            Atan22 = -PIh
        End If
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
Public Function DivRem(ByVal a As Long, ByVal b As Long, ByRef remainder As Long) As Long
    DivRem = a \ b
    remainder = a - (b * DivRem)                 ' this is about 2x faster than Mod.
End Function

Public Function LogBase(ByVal d As Double, ByVal NewBase As Double) As Double
    LogBase = Log(d) / Log(NewBase)
End Function
