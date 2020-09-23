Attribute VB_Name = "Macro_Commands_Module"
  Option Explicit

' Special synthetic macro-commands to simplify working with the
' list box called (Work), used to display all computations.


' PRINT (Output_Expression) to next available (Work) line.  It
' works very much like the VB PRINT command except that it only
' applies to the (Work) display.

  Public Sub OUT(Expression)
  Form1.Work.AddItem Expression
  End Sub

' Print a BLANK LINE to the next available (Work) line.  It works
' exactly like two BR commands in succession in HTML.
  Public Sub BLL()
  Form1.Work.AddItem " "
  End Sub
