Attribute VB_Name = "Geocentric_XYZ"
  Option Explicit

' Compute dynamical geometric gXYZ for planet at given JDE

  Public Function gXYZ_For(Planet_Name, At_JDE)
' Level 02
' DEPENDENCY: 01 hLBR_For()
'             00 Error_In()
'             00 L_Val()
'             00 B_Val()
'             00 R_Val()
'             00 Sine()
'             00 Cosine()

' Heliocentric LBR coordinate vectors for planet and Earth
  Dim hLBRp_Vector As String
  Dim hLBRe_Vector As String

' Heliocentric planet LBR and XYZ coordinate values
  Dim Lp As Double
  Dim Bp As Double
  Dim Rp As Double
  Dim Xp As Double
  Dim Yp As Double
  Dim Zp As Double

' Heliocentric Earth LBR coordinate values
  Dim Le As Double
  Dim Be As Double
  Dim Re As Double

' Geocentric planet XYZ coordinate values
  Dim gX As Double
  Dim gY As Double
  Dim gZ As Double
      
' Compute heliocentric LBR vector for planet
  hLBRp_Vector = hLBR_For(Planet_Name, At_JDE)

' Check for returned error in hLBRp_Vector
  If Error_In(hLBRp_Vector) Then _
     gXYZ_For = hLBRp_Vector: Exit Function

' Compute heliocentric LBR vector for Earth
  hLBRe_Vector = hLBR_For("Earth", At_JDE)
  
' Get heliocentric planet L coordinate value
  Lp = L_Val(hLBRp_Vector)

' Get heliocentric planet B coordinate value
  Bp = B_Val(hLBRp_Vector)

' Get heliocentric planet R coordinate value
  Rp = R_Val(hLBRp_Vector)

' Get heliocentric Earth L coordinate value
  Le = L_Val(hLBRe_Vector)

' Get heliocentric Earth B coordinate value
  Be = B_Val(hLBRe_Vector)

' Get heliocentric Earth R coordinate value
  Re = R_Val(hLBRe_Vector)

' At this point, if selected planet = sun, then set Re = 0
  If UCase(Trim(Planet_Name)) = "SUN" Then Re = 0
 
' Compute geocentric planet X coordinate value
  gX = Rp * Cosine(Bp) * Cosine(Lp) - Re * Cosine(Be) * Cosine(Le)

' Compute geocentric planet Y coordinate value
  gY = Rp * Cosine(Bp) * Sine(Lp) - Re * Cosine(Be) * Sine(Le)

' Compute geocentric planet Z coordinate value
  gZ = Rp * Sine(Bp) - Re * Sine(Be)

' Return the computed geoocentric XYZ coordinates
' of planet as delimited data vector.
  gXYZ_For = Trim(gX) & "|" & Trim(gY) & "|" & Trim(gZ)

  End Function


