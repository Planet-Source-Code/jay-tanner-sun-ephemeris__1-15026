Attribute VB_Name = "Ecliptical_Coords_From_Equatorial"
  Option Explicit

' Convert geocentric equatorial coordinates into ecliptical
' longitude and latitude angles.  All angles are expressed
' in degrees.  Coordinates are returned as a two dimensional
' delimited vector in "Ecl_Lng|Ecl_Lat" format.

  Public Function ECL_Coords_From(RA_Ang, Decl_Ang, Ecl_Obl)
' Level 1
' DEPENDENCY:  00 Sine()
'              00 Cosine()
'              00 ArcSin()
'              00 ArcTan2()

  Dim Ecl_Lng As Double
  Dim Ecl_Lat As Double

    Ecl_Lng = ArcTan2((Sine(Ecl_Lng) * Cosine(Ecl_Obl) _
       + Tangent(Ecl_Lat) * Sine(Ecl_Obl)), Cosine(Ecl_Lng))

  Ecl_Lat = ArcSin(Sine(Ecl_Lat) * Cosine(Ecl_Obl) _
       - Cosine(Ecl_Lat) * Sine(Ecl_Obl) * Sine(Ecl_Lng))

  ECL_Coords_From = Trim(Ecl_Lng) & "|" & Trim(Ecl_Lat)
  
  End Function
 


