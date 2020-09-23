Attribute VB_Name = "Heliocentric_LBR"
  Option Explicit

' Compute dynamical geometric hLBR for planet at given JDE
  
  Public Function hLBR_For(Planet_Name, At_JDE)
' Level 01
' DEPENDENCY:  00 hLBR_For_Sun()


  Dim hLBR As String
  Dim PN As String
      PN = UCase(Trim(Planet_Name))

' Compute hLBR for specified planet (Sun may be treated like planet)

  If PN = "SUN" Then _
     hLBR = hLBR_For_Earth(At_JDE, "S"): GoTo OK

  If PN = "EARTH" Then _
     hLBR = hLBR_For_Earth(At_JDE, "E"): GoTo OK


' Drop through here if invalid planet name
INVALID_PLANET_NAME:

  hLBR_For = "ERROR: " & Planet_Name & " = Invalid planet name."
  Exit Function

' Branch to here if planet hLBR computation was successful and
' return the computed delimited hLBR coordinate vector.
OK:
  
  hLBR_For = hLBR

  End Function




