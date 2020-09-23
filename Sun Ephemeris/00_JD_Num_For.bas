Attribute VB_Name = "JD_Number_At_00_hour"
  Option Explicit

'
' Compute Julian day number (JD) value for given date string
' at 00h on the given date.
'
' This function checks for invalid dates.
'
' In astronomy, the day begins at noon instead of at  midnight
' as on the civil calendar which means that the JD value is 12
' hours (0.5 day) behind the civil JD number value.
'
' For example, noon of 31 Dec 1996 marks the instant of the
' beginning of JD 2450449.
'
' On the civil calendar, since dates are reckoned from midnight,
' instead of noon, the actual value of JD for 0h on 31 Dec 1996
' is (2450449 - 0.5) = 2450448.5, which is the JD value that
' would be used in astronomical computations referring to 0h on
' that calendar date.
'
' The date argument has the general format:  "Dd Mmm Yyyy BC|AD"
'
' Typical valid date string examples are:
' "1 Jan 4713 BC"   or   "20 May 1066 AD"   or   "4 Jul 1776"
'
' The "BC" is optional as required.  The "AD" is always implied
' and assumed unless "BC" is specifically indicated.
'
' This function automatically selects the Julian or Gregorian
' calendar mode depending on the given date.

  Public Function JD_Num_For(Dd_Mmm_Yyyy_BCAD)
' Version 2.0
' Level 00

  Dim D    As Single ' Day
  Dim M    As String ' Month
  Dim Y    As String ' Year
  Dim DS   As String ' Full date string argument
  Dim Mmm  As String ' Month abbreviation (Jan to Dec)
  Dim Yyyy As String ' Year string with "BC|AD" suffix
  Dim G    As Single ' Julian/Gregorian mode flag
  Dim JD   As String ' JD number
  Dim W    As String ' Random work

' Auxiliary variables
  Dim Q   As String
  Dim R   As Double
  Dim S   As Double
  Dim T   As Double
  Dim U   As Double
  Dim V   As Double
  Dim i  As Integer
  Dim j  As Integer
  Dim k  As Integer
  
  Dim NumChars As String  ' Numerical ASCII characters
  Dim MAbbrevs As String  ' Month name abbreviations

' Recomputed date used to check for invalid date argument
  Dim Check_Date As String

' Define month abbreviations string
  MAbbrevs = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
 
' Define numerical character string
  NumChars = "0123456789"
 
' Read input date string argument
  DS = Trim(UCase(Dd_Mmm_Yyyy_BCAD))

' Check for negative sign
  If InStr(DS, "-") > 0 Then GoTo ERROR_HANDLER
  
' Extract numerical value of the month day from (DS).
  D = Val(DS)
     
  Check_Date = Trim(D) & " "
 
' Extract the three letter month abbreviation from the date
' string and determine the corresponding month number (1 to 12).
      W = ""
  For i = 1 To Len(DS)
      If InStr(NumChars, Mid(DS, i, 1)) = 0 Then Exit For
  Next i
       W = Trim(Mid(DS, i, Len(DS)))
       M = Trim(Left(W, 1) & Mid(W, 2, 2))
           Check_Date = Check_Date & M & " "
       M = 1 + Int(InStr(1, MAbbrevs, M) - 1) / 3
               
' Extract value of the year from the date string and normalize
' the numerical value for BC era if required.
  For i = 1 To Len(W)
      If InStr(NumChars, Mid(W, i, 1)) <> 0 Then Exit For
  Next i
       Y = Trim(Mid(W, i, Len(W)))
           If Right(Y, 2) <> "BC" Then Y = Val(Y) _
           Else Y = 1 - Val(Y)

 If Y <= 0 Then
    Check_Date = Check_Date & (1 - Y) & " BC"
 Else
    Check_Date = Check_Date & Y & " AD"
 End If

' At this point, the three numerical date variables, D, M and Y,
' should now be ready for use in the subsequent JD computation.

' First compute the JD number according to the old Julian calendar.
  k = Int((14 - M) / 12)
 JD = D + Int(367 * (M + (k * 12) - 2) / 12) _
    + Int(1461 * (Y + 4800 - k) / 4) - 32113

' Auto-select the proper calendar mode. If the date is prior
' to 15 Oct 1582, then use the Julian calendar, otherwise
' use the Gregorian calendar.
' The official final date on the old Julian calendar was
' Thursday, 4 Oct, 1582, which was followed by the first
' official date on the Gregorian calendar, Friday, 15 Oct, 1582.
  If JD > 2299160 Then
     JD = JD - (Int(3 * Int((Y + 100 - k) / 100) / 4) - 2)
  End If

  GoSub INV_JD

  If Q <> Check_Date Then GoTo ERROR_HANDLER
  
' Done - Return the astronomical JD value for given
' date and time of day.
  JD_Num_For = JD - 0.5

  Exit Function

INV_JD:

  If JD < 2299161 Then G = 0 Else G = 1

' Compute auxiliary values
  Q = G * Int((JD / 36524.25) - 51.12264)
  R = JD + G + Q - Int(Q / 4)
  S = R + 1524
  T = Int((S / 365.25) - 0.3343)
  U = Int(T * 365.25)
  V = Int((S - U) / 30.61)

' Compute the raw, numerical calendar date elements
  D = S - U - Int(V * 30.61)
  M = (V - 1) + 12 * (V > 13.5)
  Y = T - (M < 2.5) - 4716

' At this point the raw numerical values of D, M and Y
' have been computed.  Now they must be converted into
' the standard date format, "Dd Mmm Yyyy BC|AD", for
' output.

' Day of the month (1 to 31)
  D = Trim(D)

' Determine English month abbreviation (Jan to Dec)
  Mmm = " " & _
  Mid(MAbbrevs, 3 * (M - 1) + 1, 3)
  Mmm = Mmm & " "

' Determine the year in BC|AD format
  If Y < 0 Then
     Yyyy = Trim(1 - Y) & " BC"
  Else
     Yyyy = Trim(Y) & " AD"
  End If

' Finally, return the recomputed standard date string
' in the same format as  "12 Jan 2000 BC|AD".
' If this date is different than the date argument,
' then the date argument was invalid.
  Q = D & Mmm & Yyyy

  Return

' Return error message for invalid date argument
ERROR_HANDLER:
  JD_Num_For = "ERROR: """ & Trim(Dd_Mmm_Yyyy_BCAD) _
  & """ = Invalid calendar date"
  
  End Function

