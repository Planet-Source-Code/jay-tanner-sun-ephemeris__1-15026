Attribute VB_Name = "Equatorial_Coords_From_Ecliptical"
  Option Explicit

' Convert geocentric ecliptical coordinates into equatorial
' right ascension and declination angles.  All angles are
' expressed in degrees.  Coordinates are returned as a two
' dimensional delimited vector in "RA|Decl" format.

  Public Function EQU_Coords_From(Ecl_Lng, Ecl_Lat, Ecl_Obl)
' Level 1
' DEPENDENCY:  00 Sine()
'              00 Cosine()
'              00 ArcSin()
'              00 ArcTan2()

  Dim RA   As Double
  Dim Decl As Double

    RA = ArcTan2((Sine(Ecl_Lng) * Cosine(Ecl_Obl) _
       - Tangent(Ecl_Lat) * Sine(Ecl_Obl)), Cosine(Ecl_Lng))

  Decl = ArcSin(Sine(Ecl_Lat) * Cosine(Ecl_Obl) _
       + Cosine(Ecl_Lat) * Sine(Ecl_Obl) * Sine(Ecl_Lng))

  EQU_Coords_From = Trim(RA) & "|" & Trim(Decl)
  
  End Function


