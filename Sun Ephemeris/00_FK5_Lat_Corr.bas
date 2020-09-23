Attribute VB_Name = "FK5_Latitude_Correction"
  Option Explicit

' Compute the correction required to convert VSOP87 dynamical
' latitude into the FK5 system latitude.

' The argument (Ecl_Lng) is in decimal degrees and so is
' the returned correction value.

  Public Function FK5_Lat_Corr(At_JDE, Ecl_Lng)
' LEVEL 0

  Dim T As Double
      T = (At_JDE - 2451545) / 36525
      
  Dim Lprime As Double
  
  Lprime = (Ecl_Lng - 1.397 * T - 0.00031 * T * T) * Atn(1) / 45

  FK5_Lat_Corr = (0.03916 * (Cos(Lprime) - Sin(Lprime))) / 3600
  
  End Function
  
