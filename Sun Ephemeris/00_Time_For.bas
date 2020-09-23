Attribute VB_Name = "Time_String_For_Seconds"
  Option Explicit

' Given a time as seconds, convert to time in the standard
' format of  01:02:34.5678
'
' The result may be rounded to a maximum 4 decimals.

  Public Function Time_For(Seconds, Decimals)
' Level 00

' Hours, Minutes, Seconds
  Dim HH As String
  Dim MM As String
  Dim SS As String

  Dim Q  As String ' Random work
  Dim S  As Double ' Input argument value

' Initialize format control string
  Dim F$
      F$ = "0#"

  Dim Sign As String

' Read and adjust number of decimals to display (0 to 4).
' Limit to maximum of 4 decimals.
  Dim D As Single
      D = Abs(Decimals)
      If D > 4 Then D = 4

' Create format string corresponding to (Decimals)
  If D = 1 Then F$ = F$ & ".0"
  If D > 1 Then F$ = F$ & "." & String(D - 1, "#") & "0"
    
  S = Val(Seconds) ' Read the seconds argument

' Account for sign of argument
  If S >= 0 Then Sign = "" Else Sign = "-": S = -S
  
' Compute hours
  HH = Int(S / 3600): S = S - 3600 * HH

' Compute minutes
  MM = Int(S / 60): S = S - 60 * MM

' Compute seconds
  SS = Format(S, F$)

' Correct for any values of 60
  If Val(SS) = 60 Then MM = MM + 1: SS = ""
  If MM = 60 Then HH = HH + 1: MM = ""
  
' Format and output the equivalent time string
  If HH = 0 Then HH = "00:" Else HH = Format(HH, "0#") & ":"
  If MM <> "" Then MM = Format(MM, "0#") & ":"
  If SS <> "" Then SS = SS
  Time_For = Trim(Sign & HH & MM & SS)
  
  End Function


