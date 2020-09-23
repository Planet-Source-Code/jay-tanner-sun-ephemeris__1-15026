Attribute VB_Name = "Nutation_In_Obliquity"
  Option Explicit

' This function computes the nutation in obliquity of the ecliptic.
' It is also included within the Ecliptic() function.

' Compute nutational correction for obliquity of the ecliptic in degrees.
' In terms of accuracy, the value in arc seconds is to about ±0.001"

' This correction is applied to the mean obliquity to obtain the true or
' apparent obliquity at any given moment.

' This computation is based on the "1980 IAU Theory of Nutation" and
' includes only the terms with coefficients > 0.0003 arcsecond.

  Public Function Delta_Epsilon(At_JDE)
' LEVEL 0

  Dim JD As String ' JD number for date and time
  Dim T  As Double ' Julian centuries since J2000.0
  Dim T2 As Double ' T to the power of 2
  Dim T3 As Double ' T to the power of 3
  
  Dim Q  As Double  ' Nutation series accumulator
  
  Dim V  As Double  ' Mean elongation of the moon from the sun
  Dim W  As Double  ' Mean anomaly of the sun
  Dim X  As Double  ' Mean anomaly of the moon
  Dim Y  As Double  ' Moon's argument of latitude
  
' Longitude of ascending node of lunar orbit on the ecliptic
' as measured from the mean equinox of date.
  Dim Z  As Double
    
  T = (At_JDE - 2451545#) / 36525
   
' Compute the mean elongation of the moon in radians
  V = 297.85036 + 445267.11148 * T - 0.0019142 * T2 + T3 / 189474
  V = V * Atn(1) / 45
  
' Compute the mean anomaly of the sun in radians
  W = 357.52772 + 35999.05034 * T - 0.0001603 * T2 - T3 / 300000
  W = W * Atn(1) / 45
  
' Compute the mean anomaly of the moon in radians
  X = 134.96298 + 477198.867398 * T + 0.0086972 * T2 + T3 / 56250
  X = X * Atn(1) / 45
  
' Compute the moon's argument of latitude in radians
  Y = 93.27191 + 483202.017538 * T - 0.0036825 * T2 + T3 / 327270
  Y = Y * Atn(1) / 45
  
' Compute the longitude of moon's ascending node in radians
  Z = 125.04452 - 1934.136261 * T + 0.0020708 * T2 + T3 / 450000
  Z = Z * Atn(1) / 45
  
' Proceed to compute the nutation in obliquity
  Q = Cos(Z) * (92025 + 8.9 * T)
  Q = Q + Cos(2 * (Y - V + Z)) * (5736 - 3.1 * T)
  Q = Q + Cos(2 * (Y + Z)) * (977 - 0.5 * T)
  Q = Q + Cos(2 * Z) * (0.5 * T - 895)
  Q = Q + Cos(W) * (54 - 0.1 * T)
  Q = Q - 7 * Cos(X)
  Q = Q + Cos(W + 2 * (Y - V + Z)) * (224 - 0.6 * T)
  Q = Q + 200 * Cos(2 * Y + Z)
  Q = Q + Cos(X + 2 * (Y + Z)) * (129 - 0.1 * T)
  Q = Q + Cos(2 * (Y - V + Z) - W) * (0.3 * T - 95)
  Q = Q - 70 * Cos(2 * (Y - V) + Z)
  Q = Q - 53 * Cos(2 * (Y + Z) - X)
  Q = Q - 33 * Cos(X + Z)
  Q = Q + 26 * Cos(2 * (V + Y + Z) - X)
  Q = Q + 32 * Cos(Z - X)
  Q = Q + 27 * Cos(X + 2 * Y + Z)
  Q = Q - 24 * Cos(2 * (Y - X) + Z)
  Q = Q + 16 * Cos(2 * (V + Y + Z))
  Q = Q + 13 * Cos(2 * (X + Y + Z))
  Q = Q - 12 * Cos(X + 2 * (Y - V + Z))
  Q = Q - 10 * Cos(2 * Y + Z - X)
  Q = Q - 8 * Cos(2 * V - X + Z)
  Q = Q + 7 * Cos(2 * (W - V + Y + Z))
  Q = Q + 9 * Cos(W + Z)
  Q = Q + 7 * Cos(X + Z - 2 * V)
  Q = Q + 6 * Cos(Z - W)
  Q = Q + 5 * Cos(2 * (V + Y) - X + Z)
  Q = Q + 3 * Cos(X + 2 * (Y + V + Z))
  Q = Q - 3 * Cos(W + 2 * (Y + Z))
  Q = Q + 3 * Cos(2 * (Y + Z) - W)
  Q = Q + 3 * Cos(2 * (V + Y) + Z)
  Q = Q - 3 * Cos(2 * (X + Y + Z - V))
  Q = Q - 3 * Cos(X + 2 * (Y - V) + Z)
  Q = Q + 3 * Cos(2 * (V - X) + Z)
  Q = Q + 3 * Cos(2 * V + Z)
  Q = Q + 3 * Cos(2 * (Y - V) + Z - W)
  Q = Q + 3 * Cos(Z - 2 * V)
  Q = Q + 3 * Cos(2 * (X + Y) + Z)

' Return result in decimal degrees
  Delta_Epsilon = Q / 36000000

  End Function


