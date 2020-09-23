Attribute VB_Name = "Nutation_In_Longitude"
  Option Explicit

' Nutational correction in ecliptical longitude

' Compute nutation in ecliptical longitude in decimal degrees.  In terms
' of accuracy, the value in arc seconds is to about ±0.001"
'
  Public Function Delta_Psi(At_JDE)
' LEVEL 0

  Dim T  As Double ' Julian centuries since J2000.0
  Dim T2 As Double ' T to the power of 2
  Dim T3 As Double ' T to the power of 3
  
  Dim Q As Double  ' Nutation series accumulator
  
  Dim V As Double  ' Mean elongation of the moon from the sun
  Dim W As Double  ' Mean anomaly of the sun
  Dim X As Double  ' Mean anomaly of the moon
  Dim Y As Double  ' Moon's argument of latitude
  
' Longitude of ascending node of lunar orbit on the ecliptic
' measured from the mean equinox of date.
  Dim Z As Double
  
  T = (At_JDE - 2451545#) / 36525
   
' Compute Mean elongation of moon in radians
  V = 297.85036 + 445267.11148 * T - 0.0019142 * T2 + T3 / 189474
  V = V * Atn(1) / 45
  
' Compute mean anomaly of the sun in radians
  W = 357.52772 + 35999.05034 * T - 0.0001603 * T2 - T3 / 300000
  W = W * Atn(1) / 45
  
' Compute mean anomaly of moon in radians
  X = 134.96298 + 477198.867398 * T + 0.0086972 * T2 + T3 / 56250
  X = X * Atn(1) / 45
  
' Compute moon's argument of latitude in radians
  Y = 93.27191 + 483202.017538 * T - 0.0036825 * T2 + T3 / 327270
  Y = Y * Atn(1) / 45
  
' Compute longitude of moon's ascending node in radians
  Z = 125.04452 - 1934.136261 * T + 0.0020708 * T2 + T3 / 450000
  Z = Z * Atn(1) / 45
  
' Proceed to compute the nutation in longitude in arc seconds
  Q = Sin(Z) * (-174.2 * T - 171996)
  Q = Q + Sin(2 * (Y + Z - V)) * (-1.6 * T - 13187)
  Q = Q + Sin(2 * (Y + Z)) * (-2274 - 0.2 * T)
  Q = Q + Sin(2 * Z) * (0.2 * T + 2062)
  Q = Q + Sin(W) * (1426 - 3.4 * T)
  Q = Q + Sin(X) * (0.1 * T + 712)
  Q = Q + Sin(2 * (Y + Z - V) + W) * (1.2 * T - 517)
  Q = Q + Sin(2 * Y + Z) * (-0.4 * T - 386)
  Q = Q - 301 * Sin(2 * (Y + Z) + X)
  Q = Q + Sin(2 * (Y + Z - V) - W) * (217 - 0.5 * T)
  Q = Q - 158 * Sin(X - 2 * V)
  Q = Q + Sin(2 * (Y - V) + Z) * (129 + 0.1 * T)
  Q = Q + 123 * Sin(2 * (Y + Z) - X)
  Q = Q + 63 * Sin(2 * V)
  Q = Q + Sin(X + Z) * (0.1 * T + 63)
  Q = Q - 59 * Sin(2 * (V + Y + Z) - X)
  Q = Q + Sin(Z - X) * (-0.1 * T - 58)
  Q = Q - 51 * Sin(2 * Y + X + Z)
  Q = Q + 48 * Sin(2 * (X - V))
  Q = Q + 46 * Sin(2 * (Y - X) + Z)
  Q = Q - 38 * Sin(2 * (V + Y + Z))
  Q = Q - 31 * Sin(2 * (X + Y + Z))
  Q = Q + 29 * Sin(2 * X)
  Q = Q + 29 * Sin(2 * (Y + Z - V) + X)
  Q = Q + 26 * Sin(2 * Y)
  Q = Q - 22 * Sin(2 * (Y - V))
  Q = Q + 21 * Sin(2 * Y + Z - X)
  Q = Q + Sin(2 * W) * (17 - 0.1 * T)
  Q = Q + 16 * Sin(2 * V - X + Z)
  Q = Q + Sin(2 * (W + Y + Z - V)) * (0.1 * T - 16)
  Q = Q - 15 * Sin(W + Z)
  Q = Q - 13 * Sin(X + Z - 2 * V)
  Q = Q - 12 * Sin(Z - W)
  Q = Q + 11 * Sin(2 * (X - Y))
  Q = Q - 10 * Sin(2 * (Y + V) + Z - X)
  Q = Q - 8 * Sin(2 * (Y + V + Z) + X)
  Q = Q + 7 * Sin(2 * (Y + Z) + W)
  Q = Q - 7 * Sin(X - 2 * V + W)
  Q = Q - 7 * Sin(2 * (Y + Z) - W)
  Q = Q - 7 * Sin(2 * V + 2 * Y + Z)
  Q = Q + 6 * Sin(2 * V + X)
  Q = Q + 6 * Sin(2 * (X + Y + Z - V))
  Q = Q + 6 * Sin(2 * (Y - V) + X + Z)
  Q = Q - 6 * Sin(2 * (V - X) + Z)
  Q = Q - 6 * Sin(2 * V + Z)
  Q = Q + 5 * Sin(X - W)
  Q = Q - 5 * Sin(2 * (Y - V) + Z - W)
  Q = Q - 5 * Sin(Z - 2 * V)
  Q = Q - 5 * Sin(2 * (X + Y) + Z)
  Q = Q + 4 * Sin(2 * (X - V) + Z)
  Q = Q + 4 * Sin(2 * (Y - V) + W + Z)
  Q = Q + 4 * Sin(X - 2 * Y)
  Q = Q - 4 * Sin(X - V)
  Q = Q - 4 * Sin(W - 2 * V)
  Q = Q - 4 * Sin(V)
  Q = Q + 3 * Sin(2 * Y + X)
  Q = Q - 3 * Sin(2 * (Y + Z - X))
  Q = Q - 3 * Sin(X - V - W)
  Q = Q - 3 * Sin(W + X)
  Q = Q - 3 * Sin(2 * (Y + Z) + X - W)
  Q = Q - 3 * Sin(2 * (V + Y + Z) - W - X)
  Q = Q - 3 * Sin(2 * (Y + Z) + 3 * X)
  Q = Q - 3 * Sin(2 * (V + Y + Z) - W)

' Return result in degrees
  Delta_Psi = Q / 36000000

  End Function

