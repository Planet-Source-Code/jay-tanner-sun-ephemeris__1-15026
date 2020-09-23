Attribute VB_Name = "Angle_Output"
  Option Explicit

' Astronomical angle output function.
' An angle argument given in decimal degrees may be output in any of
' several different units.  The angle can be an hour angle or any
' longitude or latitude angle in general.
'
' The purpose of this function is to help produce a uniform output
' that may easily be arranged into neat, tabular columns.
'
' Output codes for the output modes are:
' DD  = Decimal degrees - Decimals = 11
' DMS = Degrees, Minutes and Seconds - Decimals = 2
' DH  = Decimal hours - Decimals = 11
' HMS = Hours, Minutes and Seconds - Decimals = 3
'
' Decimals = Number of decimal places used for output.
' All formatted strings are padded on left to make 16 characters.

  Public Function Ang_Out(Degrees_In, Out_Mode, Pos_Sign As Boolean)
' Level 0

  Dim Q  As Variant

  Dim A  As Double
      A = Val(Degrees_In)

  Dim Sign As String

  If A >= 0 Then Sign = "+" Else Sign = "-"
  If Pos_Sign = False And A >= 0 Then Sign = ""
  A = Abs(A)

  Dim U As String
      U = UCase(Trim(Out_Mode))
        
  Dim dd  As String
  Dim HH  As String
  Dim MM  As String
  Dim SS  As String
  Dim S   As Double
  
  If U = "DD" Then
     Q = Sign & Format(A, "#0.##########0")
     Ang_Out = Right(Space(16) & Q & "°", 16)
  Exit Function
  End If

  If U = "DMS" Then GoTo DMS_OUT

' Handle decimal hours
  If U = "DH" Then
     Q = Sign & Format(A / 15, "#0.##########0")
     Ang_Out = Right(Space(16) & Q & "h", 16)
     Exit Function
  End If

  If U = "HMS" Then GoTo HMS_OUT

' Drop through here if invalid mode
  Ang_Out = "ERROR: """ & Out_Mode & """ = Invalid angle output mode"
  Exit Function
 
DMS_OUT:
  S = A * 3600 ' Convert angle from degrees into arc seconds
  
' Compute degrees
  dd = Int(S / 3600): S = S - 3600 * dd
' Compute minutes
  MM = Int(S / 60): S = S - 60 * MM
' Compute seconds
  SS = S
  
' Correct for any values of 60
  If Val(SS) = 60 Then MM = MM + 1: SS = ""
  If MM = 60 Then dd = dd + 1: MM = ""
  
' Format the angle in DMS format
  dd = Format(dd, "#0") & "° "
  MM = Format(MM, "0#") & "' "
  SS = Format(SS, "0#.#0") & """"

' Return the computed angle
  Q = Trim(Sign & dd & MM & SS)
  Ang_Out = Right(Space(16) & Q, 16)
  
  Exit Function

HMS_OUT:
  S = A * 240 ' Convert hour angle from degrees into seconds
  
' Compute hours
  HH = Int(S / 3600): S = S - 3600 * HH
' Compute minutes
  MM = Int(S / 60): S = S - 60 * MM
' Compute seconds
  SS = S
  
' Correct for any values of 60
  If Val(SS) = 60 Then MM = MM + 1: SS = ""
  If MM = 60 Then HH = HH + 1: MM = ""
  
' Format the hour angle in HMS format
  HH = Format(HH, "0#") & "h "
  MM = Format(MM, "0#") & "m "
  SS = Format(SS, "0#.##0") & "s"

' Return the computed hour angle
  Q = Trim(Sign & HH & MM & SS)
  Ang_Out = Right(Space(16) & Q, 16)
   
  End Function


