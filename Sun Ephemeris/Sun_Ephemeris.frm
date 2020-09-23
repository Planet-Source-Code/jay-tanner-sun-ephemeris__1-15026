VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Sun - Geocentric Ephemeris Generator v1.5 - NeoProgrammics"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "Sun_Ephemeris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton About_Button 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9675
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   " Some Info "
      Top             =   315
      Width           =   195
   End
   Begin VB.TextBox The_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2070
      TabIndex        =   1
      Text            =   "00:00:00"
      Top             =   270
      Width           =   1320
   End
   Begin VB.TextBox The_Date 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Text            =   "Jan 2001"
      Top             =   270
      Width           =   1950
   End
   Begin MSComDlg.CommonDialog SAVE_Dialog 
      Left            =   7515
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   " Save Ephemeris to Text File "
      Filter          =   "Text File|*.txt"
   End
   Begin VB.TextBox The_Delta_T 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3510
      TabIndex        =   2
      Text            =   "00:00:00"
      Top             =   270
      Width           =   1770
   End
   Begin VB.ListBox Work 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7080
      Left            =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   630
      Width           =   9825
   End
   Begin VB.CommandButton Compute_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Compute"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   " Compute the Indicated Ephemeris "
      Top             =   270
      Width           =   1005
   End
   Begin VB.CommandButton SAVE_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Ephemeris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8055
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   " Save Ephemeris as a Text File "
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delta T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3510
      TabIndex        =   8
      ToolTipText     =   " UTC  +  Delta T  =  Dynamical Time "
      Top             =   45
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Universal Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2115
      TabIndex        =   7
      ToolTipText     =   " UTC  =  Dynamical Time - Delta T "
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Month and Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   6
      ToolTipText     =   " Month and Year in Same Format as:  Jan 2001 "
      Top             =   45
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

' ================================================================
' Define computed ephemeris data vector array. The results of each
' element of a computation is held in this array for easy access later.
  Public Ephemeris_Line As Variant

' =============================================================
' Define an error status flag to indicate last operation caused
' an error.
  Public ERROR_Flag_Set As Boolean

' ==========================================
' What to do when this program is terminated
  Private Sub Form_Terminate()

  Unload Me

  End Sub

' ========================================================
' Unhighlight any selected line in work area when a blank
' area of the form is clicked on.
  Private Sub Form_Click()
  Work.ListIndex = -1
  Compute_Button.SetFocus
  End Sub

' ==================================
  Private Sub Compute_Button_Click()

' Random work variables
  Dim Q
  Dim W

' Reset error status flag
  ERROR_Flag_Set = False

' Unhighlight any selected line in the work area.
  Work.ListIndex = -1

' Loop cycle control counters
  Dim D  As Integer
  Dim i  As Integer
  Dim j  As Integer

' Formatted string copy of D value
  Dim dd As String

' JD and time values used in computations
  Dim JD       As String
  Dim JDE      As String
  Dim TimeFrac As String

' Reformat the time input argument just for neatness.
  Q = Day_Frac_For(The_Time): Q = 86400 * Q
  Q = Time_For(Q, 0): The_Time = Q

' Also reformat the delta T input argument just for neatness.
  Q = Day_Frac_For(The_Delta_T): Q = 86400 * Q
  Q = Time_For(Q, 1): The_Delta_T = Q

' Print ephemeris heading
  Work.Clear
  OUT " Geocentric Ephemeris of Sun for the Month of " & The_Date
  OUT " at " & The_Time & " UTC on Each Date      Delta T = " _
    & The_Delta_T

  OUT String(80, "-")
  OUT " Day |       RA        |       Decl      |" _
    & "    Distance    | Semi Diam"

' Change mouse pointer to hourglass while computing
  Form1.MousePointer = vbHourglass

' Compute raw geocentric computations output vector for each
' date of the given month at the specified time on each date.
      j = 0
  For D = 1 To 31
  dd = Str(D): If D < 10 Then dd = " " & dd
  dd = dd & "  |"

' Compute JD number for 0h on date
  JD = JD_Num_For(Str(D) & The_Date)
       If Error_In(JD) Then
          Form1.MousePointer = vbDefault ' Restore normal mouse pointer
          Work.ListIndex = -1
          If j = 1 Then Exit Sub
          Work.Clear
          OUT " "
          OUT " ERROR: Invalid starting month & year."
          OUT " "
          OUT " Valid format = Oct 1956 BC|AD     (AD = Optional)"
          Beep
          ERROR_Flag_Set = True
          Exit Sub
       End If

' Compute and add fraction corresponding to the UTC
' and the Delta T value, if known, to obtain the
' complete dynamical ephemeris time value to be used
' for the computation.
  JDE = JD + Day_Frac_For(The_Time) _
      + Day_Frac_For(The_Delta_T)

  Q = gLBR_For("Sun", JDE, "EQU")

' Parse and display computed geocentric position of planet
  Ephemeris_Line = Parsed(Q)
  Q = Output_Computed(Ephemeris_Line)

  OUT dd & Q

  j = 1

  Next D

' Change mouse pointer back to normal after computing
  Form1.MousePointer = vbDefault
  Work.ListIndex = -1

  End Sub

' =====================================================
' Function to construct a single ephemeris output line
' from a returned computation held in the public array.

  Private Function Output_Computed(Ephemeris_Line)

  Dim Q     As Variant
  Dim RA    As Variant
  Dim Decl  As Variant
  Dim Dist  As Variant
  Dim vMag  As Variant
  Dim SDiam As Variant
  Dim VFrac As Variant

  RA = Ang_Out(Ephemeris_Line(0), "HMS", False)
  Decl = Ang_Out(Ephemeris_Line(1), "DMS", True)
  Dist = Dist_Out(Ephemeris_Line(2), "AU")
  SDiam = Ephemeris_Line(3)

  Output_Computed = RA & " |" & Decl & " |" & Dist & " | " _
  & SDiam

  End Function

' ===============================================================
' Modified VB6 split function to parse the returned computational
' data vector into its individual elements.

  Public Function Parsed(ByVal Data_Vector As String, _
  Optional ByVal Delimiter As String = "|", _
  Optional ByVal Limit As Long = -1, _
  Optional Compare As VbCompareMethod = vbBinaryCompare) _
  As Variant

  Dim Element As Variant
  Dim i       As Long

' Parse the individual data vector elements into an array
  Element = Split(Data_Vector, Delimiter, Limit, Compare)
  For i = LBound(Element) To UBound(Element)
          If Len(Element(i)) = 0 Then Element(i) = Delimiter
  Next i

' Returned the parsed data vector elements array
  Parsed = Filter(Element, Delimiter, False)
    
  End Function

' ====================================================
' Save current contents of work display to a text file

  Private Sub SAVE_Button_Click()

  On Error GoTo ERROR_HANDLER

  Dim File_Name As String  ' Name of file to save
  Dim i         As Integer ' Loop control index
  Dim Q As String

' Unhighlight any selected line in the work area.
  Work.ListIndex = -1

  Compute_Button.SetFocus

' Check error flag set status and don't save if set
  If ERROR_Flag_Set Then
  Q = MsgBox("There was a computation error." & vbCrLf _
      & "There is nothing to be saved.", vbExclamation, _
       " NeoProgrammics Sun Ephemeris")
         Compute_Button.SetFocus
         Exit Sub
  End If

' Point to current app path directory for saving ephemerides.
' Initial default save path is to the same directory
' where the program resides.
  ChDrive (App.Path)
  ChDir (App.Path)
  
  Work.Enabled = True
      
  If Work.ListCount = 0 Then
     Q = MsgBox("The computations work area is blank." & vbCrLf _
       & "There is nothing to be saved.", vbExclamation, _
        " NeoProgrammics Sun Ephemeris")
         Compute_Button.SetFocus
         Exit Sub
  End If

  SAVE_Dialog.FileName = "Sun Ephemeris for " & The_Date
  
  File_Name = ""
  SAVE_Dialog.ShowSave

  File_Name = SAVE_Dialog.FileTitle
  If File_Name = "" Then
     Compute_Button.SetFocus
     Exit Sub
  End If

  Open File_Name For Output As #1
  For i = 0 To Work.ListCount - 1
      Print #1, Work.List(i)
  Next i
  Print #1, " "

  Print #1, " Computed by NeoProgrammics Ephemeris Generator v1.5"
  Close 1

  Q = MsgBox("The work area has been saved as" & vbCrLf & vbCrLf _
    & """" & File_Name & """", vbInformation, " NeoProgrammics Sun Ephemeris")
  Compute_Button.SetFocus
  Exit Sub

' Exit if CANCEL generates an error
ERROR_HANDLER:
  Compute_Button.SetFocus
  Exit Sub

  End Sub


  Private Sub About_Button_Click()

  Dim Q   As String
  Dim CR  As String
      CR = vbCrLf
  Dim CR2 As String
      CR2 = CR & CR

' Unhighlight any selected line in the work area.
  Work.ListIndex = -1

  Compute_Button.SetFocus

' Display a little info about this program
  Q = ""
  Q = Q & "This program is based on the VSOP87 theory of planetary orbits" & CR
  Q = Q & "in spherical variables.  This theory was first published in 1987" & CR
  Q = Q & "by Pierre Bretagnon of the Bureau des Longitudes in Paris." & CR2
  Q = Q & "The full theory is implemented in this program and theoretically" & CR
  Q = Q & "is accurate to within an arcsecond or better over the range from" & CR
  Q = Q & "2000 BC to 6000 AD." & CR2
  Q = Q & "This theory is often used to compute tables of planetary events" & CR
  Q = Q & "over periods of thousands of years."



  Q = MsgBox(Q, vbInformation, _
    " Sun - Geocentric Ephemeris Generator - Written by Jay Tanner")
 




  End Sub
