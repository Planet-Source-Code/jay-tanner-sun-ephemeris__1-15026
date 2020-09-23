Attribute VB_Name = "Math_Module"
  Option Explicit

' Custom mathematical functions

' Base 10 logarithm function for positive arguments
  Public Function Log10(ArgX)
' Level 00
  Log10 = Log(ArgX) / 2.30258509299405
  End Function

  Public Function Rad(Degrees_Arg)
' Level 00
  Rad = Degrees_Arg * Atn(1) / 45
  End Function

  Public Function Deg(Radians_Arg)
' Level 00
  Deg = Radians_Arg * 45 / Atn(1)
  End Function

' =========================================================================
' Custom trigonometric and inverse trigonometric functions that take or
' return degrees instead of radian values.

  Public Function Sine(Degrees_Arg)
' Level 00
  Sine = Sin(Degrees_Arg * Atn(1) / 45)
  End Function

  Public Function ArcSin(ArgX)
' Level 00
  ArcSin = 45 * (Atn(ArgX / Sqr(-ArgX * ArgX + 1))) _
         / Atn(1)
  End Function

  Public Function Cosine(Degrees_Arg)
' Level 00
  Cosine = Cos(Degrees_Arg * Atn(1) / 45)
  End Function

  Public Function ArcCos(ArgX)
' Level 00
  ArcCos = 45 * (Atn(-ArgX / Sqr(-ArgX * ArgX + 1)) _
         + 2 * Atn(1)) / Atn(1)
  End Function

  Public Function Tangent(Degrees_Arg)
' Level 00
  Tangent = Tan(Degrees_Arg * Atn(1) / 45)
  End Function

  Public Function ArcTan(ArgX)
' Level 00
  ArcTan = 45 * Atn(ArgX) / Atn(1)
  End Function
  
  Public Function ArcTan2(ArgY, ArgX)
' Level 00

  Dim Q As Double
      Q = 1

' Check for special zero value cases
  If ArgY = 0 And ArgX >= 0 Then Q = 0
  If ArgY > 0 And ArgX = 0 Then Q = 90
  If ArgY = 0 And ArgX < 0 Then Q = 180
  If ArgY < 0 And ArgX = 0 Then Q = 270

' If none of the special zero cases apply, then do this
  If Q = 1 Then
     Q = 45 * Atn(ArgY / ArgX) / Atn(1)
     If ArgX < 0 Then Q = Q + 180
     If Q < 0 Then Q = Q + 360
  End If

' Output angle in degrees
  ArcTan2 = Q

  End Function

' =========================================================================

