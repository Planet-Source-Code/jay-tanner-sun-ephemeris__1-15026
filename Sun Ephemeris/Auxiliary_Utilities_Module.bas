Attribute VB_Name = "Auxiliary_Utilities_Module"
  Option Explicit

' Auxiliary utility functions for miscellaneous tasks
'

' Functions to extract individual LBR and XYZ values
' from a given delimited coordinate vector.
'
' The delimiter character is the bar "|"  (ANSI character 124)

  Public Function L_Val(From_LBR_Vector)
' Level 00
  L_Val = Val(Trim(From_LBR_Vector))
  End Function
  
  Public Function B_Val(From_LBR_Vector)
' Level 00
  Dim LBR As String
      LBR = Trim(From_LBR_Vector)
  B_Val = Val(Trim(Mid(LBR, InStr(LBR, "|") + 1, Len(LBR))))
  End Function

  Public Function R_Val(From_LBR_Vector)
' Level 00
  Dim LBR As String
  LBR = Trim(From_LBR_Vector)
  R_Val = Val(Trim(Mid(LBR, InStr(InStr(LBR, "|") _
          + 1, LBR, "|") + 1, Len(LBR))))
  End Function

  Public Function X_Val(From_XYZ_Vector)
' Level 00
  X_Val = Val(Trim(From_XYZ_Vector))
  End Function
  
  Public Function Y_Val(From_XYZ_Vector)
' Level 00
  Dim XYZ As String
      XYZ = Trim(From_XYZ_Vector)
  Y_Val = Val(Trim(Mid(XYZ, InStr(XYZ, "|") + 1, Len(XYZ))))
  End Function

  Public Function Z_Val(From_XYZ_Vector)
' Level 00
  Dim XYZ As String
  XYZ = Trim(From_XYZ_Vector)
  Z_Val = Val(Trim(Mid(XYZ, InStr(InStr(XYZ, "|") _
          + 1, XYZ, "|") + 1, Len(XYZ))))
  End Function

' ------------------------------------------------------------

  Public Function Error_In(Returned_Value) As Boolean
' V1.0
' Return error status of returned value of a function.
'
' This function is NOT case sensitive.
' This makes it easier to detect if an error occured within one
' of the functions.

' Just pass the returned value to this function as an argument to
' find out if an error was returned.
'
' If the returned string from a function contains the substring
' "ERROR", then return boolean "True", otherwise return "False".

  If InStr(UCase(Returned_Value), "ERROR") > 0 Then
     Error_In = True
  Else
     Error_In = False
  End If
  
  End Function




