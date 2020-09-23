Attribute VB_Name = "Geocentric_Light_Time_To_Planet"
  Option Explicit

' Compute the geocentric light-time between planet and Earth in
' Julian days at given JDE.

  Public Function Light_Time_To(Planet_Name, At_JDE)
' Level 3
' DEPENDENCY:  02 gXYZ_For()

' Geocentric rectangular coordinates
  Dim gXYZ As String
  Dim gX   As Double
  Dim gY   As Double
  Dim gZ   As Double

' Light time iteration error tolerance
  Dim ET   As Double
      ET = 1E-16

' Initialize light-time approximation to zero
  Dim LT_Approx As Double
      LT_Approx = 0

  Dim LT As Double
      LT = ET

' Start light-time iteration loop until error tolerance reached
  Do Until Abs(LT - LT_Approx) < ET
  LT_Approx = LT

' Compute geocentric rectangular coordinates for planet
  gXYZ = gXYZ_For(Planet_Name, At_JDE - LT_Approx)
  If Error_In(gXYZ) Then Light_Time_To = gXYZ: Exit Function
  gX = X_Val(gXYZ)
  gY = Y_Val(gXYZ)
  gZ = Z_Val(gXYZ)

' Compute geometric distance between planet and Earth at JDE
  LT = Sqr(gX * gX + gY * gY + gZ * gZ) * 5.77551830441213E-03

  Loop

' Return computed light time in Julian days
  Light_Time_To = LT

  End Function


