Attribute VB_Name = "Sidereal_Time"
  Option Explicit

' Compute the local mean or true apparent sidereal time angle in
' degrees for a given JDE and UT at any given longitude.
'
' Longitude arguments is in degrees according to the convention:
' East = (0 to -180) and West = (0 to +180)
' However, a longitude of 0 to 360 (westward) may be used.

  Public Function Sid_Time(At_JDE, UT, Longitude, Mean_or_Apparent)
' Level 1
' DEPENDENCY:  00 Ecliptic()
'              00 Day_Frac_For()
'              00 Delta_Psi()
'              00 Cosine()

  Dim T      As Double ' Julian centuries from J2000.0
  Dim ST     As Double ' Sidereal time angle
  Dim e      As Double ' Obliquity of the ecliptic

  Dim L      As Double

' Check if longitude given in 0 to 360 format.
' If longitude > 180 then subtract from 360 and negate result.
  L = Val(Longitude)
      If L > 180 Then L = -(360 - L)

' Check if mean or apparent sidereal time indicated
  Dim M_or_A As String
      M_or_A = Left(UCase(Trim(Mean_or_Apparent)), 1)
      If M_or_A <> "A" Then M_or_A = "M"

' Compute Julian centuries from J2000.0
  T = (At_JDE - 2451545) / 36525

' Compute obliquity of the ecliptic
  e = Ecliptic(At_JDE, M_or_A)

' Compute sidereal time at 00h at Greenwich
   ST = 100.46061837 + 36000.770053608 * T _
      + 0.000387933 * T * T _
      + T * T * T / 38710000
      
' Compute mean ST at specified longitude and UT
  ST = ST + Day_Frac_For(UT) * 360.985647366 - L
  
' Correct for nutation if true ST mode indicated, otherwise use mean value
  If M_or_A = "A" Then _
     ST = ST + Delta_Psi(At_JDE + Day_Frac_For(UT)) * Cosine(e)
      
' Modulate sidereal time angle to fall between 0 and 360 degrees
  If Abs(ST) > 360 Then ST = ST - 360 * Int(ST / 360)
  If ST < 0 Then ST = ST + 360

' Return computed ST value in degrees
  Sid_Time = ST

  End Function

