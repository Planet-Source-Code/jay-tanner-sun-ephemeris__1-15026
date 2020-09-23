Attribute VB_Name = "FK5_Longitude_Correction"
  Option Explicit

' Compute the correction required to convert VSOP87 dynamical
' ecliptical longitude into the corresponding FK5 system longitude.

' The input angle arguments and output are in decimal degrees.

  Public Function FK5_Lng_Corr(At_JDE, Ecl_Lng, Ecl_Lat)
' LEVEL 0

  Dim T As Double
      T = (At_JDE - 2451545) / 36525
      
  Dim Q As Double

  Dim Lprime As Double
  Dim B      As Double

  B = Ecl_Lat * Atn(1) / 45
  Lprime = (Ecl_Lng - 1.397 * T - 0.00031 * T * T) * Atn(1) / 45

  Q = -0.09033 + 0.03916 * (Cos(Lprime) + Sin(Lprime)) * Tan(B)
  
  FK5_Lng_Corr = Q / 3600

  End Function

