Attribute VB_Name = "Heliocentric_XYZ"
  Option Explicit

' Compute dynamical geometric hXYZ for planet at given JDE

  Public Function hXYZ_For(Planet_Name, At_JDE)
' Level 02
' DEPENDENCY: 01 hLBR_For()
'             00 Error_In()
'             00 L_Val()
'             00 B_Val()
'             00 R_Val()
'             00 Sine()
'             00 Cosine()

  Dim hLBR_Vector As String

  Dim L As Double
  Dim B As Double
  Dim R As Double
  Dim i As Integer

' Compute hLBR for planet
  hLBR_Vector = hLBR_For(Planet_Name, At_JDE)

' Check for returned error in hLBR
  If Error_In(hLBR_Vector) Then hXYZ_For = hLBR_Vector: Exit Function
  
' Get L coordinate
  L = L_Val(hLBR_Vector)

' Get B coordinate
  B = B_Val(hLBR_Vector)

' Get R coordinate
  R = R_Val(hLBR_Vector)

' Compute and return corresponding heliocentric XYZ coordinates
' as a delimited data vector.
  hXYZ_For = Trim(R * Cosine(B) * Cosine(L)) & "|" _
           & Trim(R * Cosine(B) * Sine(L)) & "|" _
           & Trim(R * Sine(B))

  End Function



