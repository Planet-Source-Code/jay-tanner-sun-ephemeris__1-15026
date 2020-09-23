Attribute VB_Name = "Geocentric_LBR"
  Option Explicit

' Compute apparent geocentric coordinates of planet at given JDE value.
' Coordinates are returned in degrees in the form of a 3D coordinate
' vector and can be in either ECLiptical or EQUatorial coordinates.

  Public Function gLBR_For(Planet_Name, At_JDE, ECL_or_EQU)
' Level 4
' DEPENDENCY:  03 Light_Time_To()
'              00 hLBR_For()
'              02 gXYZ_For()
'              00 ArcTan()
'              00 ArcTan2()
'              00 Ecliptic()
'              00 FK5_Lng_Corr()
'              00 FK5_Lat_Corr
'              00 Delta_Psi()
'              00 Error_In()
'              00 X_Val()
'              00 Y_Val()
'              00 Z_Val()
'              00 L_Val()
'              00 B_Val()
'              00 R_Val()

  Dim W As Variant ' Random work
  Dim i As Integer

' Geocentric ecliptical XYZ coordinates of planet
  Dim XYZ As String
  Dim X   As Double
  Dim Y   As Double
  Dim Z   As Double

' Geocentric ecliptical LBR coordinates of planet
  Dim L As Double
  Dim B As Double
  Dim R As Double

' Light time
  Dim LT As String

' Distance and phase related variables
  Dim d1      As Double ' Distance Earh to Sun
  Dim d2      As Double ' Distance Planet to Sun
  Dim d3      As Double ' Distance Earth to Planet
  Dim SDiam   As String   ' Semidiameter diameter

' Obliquity of the ecliptic
  Dim e As Double
      e = Ecliptic(At_JDE, "Apparent")

' Coordinate output mode flag (EQU = Default)
  Dim EC_or_EQ As String
      EC_or_EQ = Left(UCase(Trim(ECL_or_EQU)), 3)
      If EC_or_EQ <> "ECL" Then EC_or_EQ = "EQU"

' Error if planet = Earth
  If UCase(Trim(Planet_Name)) = "EARTH" Then
  gLBR_For = "ERROR: Geocentric computations do not apply to Earth."
  Exit Function
  End If

' Compute distance between Earth and Sun
  d1 = R_Val(hLBR_For("Earth", At_JDE))
      
' Compute distance between planet and Sun
  d2 = R_Val(hLBR_For(Planet_Name, At_JDE))
     
' Compute iterated light-time to planet at JDE
  LT = Light_Time_To(Planet_Name, At_JDE)
  If Error_In(LT) Then gLBR_For = LT: Exit Function

' Compute geocentric XYZ for planet and Earth at JDE
  XYZ = gXYZ_For(Planet_Name, At_JDE)
  If Error_In(XYZ) Then gLBR_For = XYZ: Exit Function

'  Get geocentric geometric XYZ coordinates of planet and
'  use then to compute the true geometric distance from
'  Earth to the planet.
   X = X_Val(XYZ)
   Y = Y_Val(XYZ)
   Z = Z_Val(XYZ)
   R = Sqr(X * X + Y * Y + Z * Z)
  d3 = R

' Compute geocentric XYZ for planet and Earth at (JDE-LT)
  XYZ = gXYZ_For(Planet_Name, At_JDE - LT)
  X = X_Val(XYZ)
  Y = Y_Val(XYZ)
  Z = Z_Val(XYZ)

' Compute geocentric ecliptical L,B values from XYZ values
  L = ArcTan2(Y, X)
  B = ArcTan(Z / Sqr(X * X + Y * Y))
    
' Apply reductions to FK5 system coordinates
  W = FK5_Lng_Corr(At_JDE, L, B)
  B = B + FK5_Lat_Corr(At_JDE, L)
  L = L + W

' Apply correction for nutation in longitude
  L = L + Delta_Psi(At_JDE)

' Determine if results are to be returned in ecliptical
' or in equatorial coordinates.
  If EC_or_EQ = "EQU" Then
     W = EQU_Coords_From(L, B, e)
     L = L_Val(W)
     B = B_Val(W)
  End If

' Construct raw 3D coordinates output vector
  gLBR_For = Trim(L) & "|" & Trim(B) & "|" & Trim(R)

' Compute angular diameter
  SDiam = 959.63 / R / 3600
  i = 10: SDiam = Right(Space(i) & Ang_Out(SDiam, "DMS", False), i)
  
' Return geocentric computations data vector consisting of:
' RA | Decl |Dist | vMag | pAng
  gLBR_For = gLBR_For & "|" & SDiam

  End Function

