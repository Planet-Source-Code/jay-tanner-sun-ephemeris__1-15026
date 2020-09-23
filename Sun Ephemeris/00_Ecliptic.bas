Attribute VB_Name = "Ecliptic_Obliquity"
  Option Explicit

' Compute the mean or apparent obliquity of the ecliptic in degrees for
' the current JDE value at any instant.
'
' This function is based on a formula derived by J. Laskar,
' Astronomy and Astrophysics, (1968) Vol. 157, page 68.
'
' The estimated accuracy of this formula is ±0.01" between the
' years 1000 AD and 3000 AD and a few arcseconds after 10000 years.
' It is only valid in the range ±10000 years either way of J2000.0
'
' Over the long term, it is more accurate than the formula adopted by
' the International Astronomical Union (IAU) for general almanac
' computations.
'
' This version has the nutational correction built in to make it a
' stand-alone, level 0 function.
'
' The second argument determines if the mean or apparent obliquity
' is returned.  Only a single letter "M" or "A" is required.

  Public Function Ecliptic(At_JDE, Mean_or_Apparent)
' LEVEL 00
 
  Dim T   As Double ' Julian time factor relative to J2000.0
  Dim T2  As Double ' T to the power of 2
  Dim T3  As Double ' T to the power of 3

  Dim Obl As Double ' Mean or apparent obliquity of the ecliptic
  Dim P   As Double ' Successive powers of T from 2 to 10
  Dim Q   As Double ' Nutational series accumulator
  
' Elements used to correct for nutation (1980 IAU Theory)
  Dim V   As Double  ' Mean elongation of the moon from the sun
  Dim W   As Double  ' Mean anomaly of the sun
  Dim X   As Double  ' Mean anomaly of the moon
  Dim Y   As Double  ' Moon's argument of latitude
  
' Longitude of ascending node of lunar orbit on the ecliptic
' as measured from the mean equinox of date.
  Dim Z   As Double

' Mean or apparent flag ("M" or "A")
  Dim M_or_A As String
      M_or_A = Left(UCase(Trim(Mean_or_Apparent)), 1)

         
  T = (At_JDE - 2451545) / 3652500

' Compute mean obliquity of the ecliptic
  Obl = 84381.448 - 4680.93 * T: P = T * T
  Obl = Obl - 1.55 * P: P = P * T
  Obl = Obl + 1999.25 * P: P = P * T
  Obl = Obl - 51.38 * P: P = P * T
  Obl = Obl - 249.67 * P: P = P * T
  Obl = Obl - 39.05 * P: P = P * T
  Obl = Obl + 7.12 * P: P = P * T
  Obl = Obl + 27.87 * P: P = P * T
  Obl = Obl + 5.79 * P: P = P * T
  Obl = Obl + 2.45 * P: P = P * T
  
  Obl = Obl / 3600 ' This is the mean obliquity in degrees

' If not apparent mode, then return mean value
  If M_or_A = "M" Then Ecliptic = Obl: Exit Function

' Drop through here, to compute the apparent ecliptic corrected
' for nutation.
      
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

' Return true apparent ecliptic obliquity in decimal degrees
  Ecliptic = Obl + Q / 36000000
   
  End Function

  
