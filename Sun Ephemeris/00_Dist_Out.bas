Attribute VB_Name = "Distance_Output"
  Option Explicit

' Function to return distance value according to interface settings.
' Kilometers and miles are displayed in millions.

  Public Function Dist_Out(AU_In, D_Units)
' Level 00

' Read the raw data value in AUs
  Dim D As String
      D = Val(AU_In)
      
' Read the distance units setings
  Dim DU As String
      DU = UCase(Trim(D_Units))

' Convert raw AUs into equivalent km or mi according to mode
  If DU = "KM" Then D = D * 149597870: GoTo KM_OUT
  If DU = "MI" Then D = D * 92955806.8380657: GoTo MI_OUT
  If DU = "AU" Then GoTo AU_OUT

' Error if invalid units
  Dist_Out = "ERROR: """ & D_Units & """ = Invalid distance units."
  Exit Function
 
AU_OUT:
     D = Format(D, "#0.#######0")
     Dist_Out = Right(Space(15) & D & " AU", 15)
     Exit Function

KM_OUT:
     D = Format(D / 1000000#, "#0.######0")
     Dist_Out = Right(Space(16) & D & " Mkm", 16)
     Exit Function
 
MI_OUT:
     D = Format(D / 1000000#, "#0.######0")
     Dist_Out = Right(Space(16) & D & " Mmi", 16)
  
  End Function

