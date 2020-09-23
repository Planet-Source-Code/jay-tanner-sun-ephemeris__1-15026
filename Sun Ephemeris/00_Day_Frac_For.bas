Attribute VB_Name = "Day_Fraction_For_Time_HHMMSS"
  Option Explicit

' =======================================================================
' Compute fraction of day corresponding to a time string argument given
' in the format  ± HH:MM:SS

  Public Function Day_Frac_For(HHMMSS)
' LEVEL 0

' Internal string pointers
  Dim i As Integer
  Dim j As Integer
  
  Dim TS  As String ' Time string in standard "HH:MM:SS" format
  Dim S   As Double ' Time expressed as seconds

  Dim Sign As Integer ' Sign value = ±1
      Sign = 1

' Read time of day argument
  TS = Trim(HHMMSS)

' Account for sign
  If Left(TS, 1) = "-" Then
     TS = Right(TS, Len(TS) - 1)
     Sign = -Sign
  End If

' If argument is null, equate it to zero
  If TS = "" Then TS = 0
    
' If no colons, then attach zero minutes and seconds
  If InStr(TS, ":") = 0 Then TS = TS & ":00:00"
     
' Mark location of 1st colon
  i = InStr(1, TS, ":")
     
' Mark location of 2nd colon
  j = InStr(i + 1, TS, ":")
  
' If no 2nd colon, then attach zero seconds
  If j = 0 Then TS = TS & ":00"
     j = InStr(i + 1, TS, ":")
  
' Parse the time values and convert into seconds
  S = Val(TS) * 3600#
  S = S + Val(Mid(TS, i + 1, Len(TS))) * 60#
  S = S + Val(Mid(TS, j + 1, Len(TS)))
     
' Return the equivalent fraction of a day equivalent to time string
  Day_Frac_For = Sign * S / 86400#
 
  End Function

