VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Create Codebook"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTextWidth 
      Height          =   330
      Left            =   720
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame frameModel 
      Caption         =   "Enigma Model"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   6855
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   12
         Scrolling       =   1
      End
      Begin VB.ComboBox cmbModel 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   5415
      End
   End
   Begin VB.Frame frameCodebook 
      Caption         =   "Codebook"
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton cmdSelect 
         Caption         =   "..."
         Height          =   300
         Left            =   6120
         TabIndex        =   6
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   3840
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "1900"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCodeName 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblFolder 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Save Folder"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   1475
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   13
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1025
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codename Net"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   650
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Me.txtCodeName.SetFocus
End Sub

Private Sub txtCodeName_GotFocus()
Me.txtCodeName.SelStart = 0
Me.txtCodeName.SelLength = Len(Me.txtCodeName.Text)
End Sub

Private Sub txtCodeName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtYear_GotFocus()
Me.txtYear.SelStart = 0
Me.txtYear.SelLength = Len(Me.txtYear.Text)
End Sub

Private Sub cmbModel_Click()
Select Case Me.cmbModel.ListIndex + 1
Case 1
    '3-rotor Wehrmacht/Luftwaffe
    Me.lblDescription.Caption = "This is the basic 3 rotor Enigma machine, used by Wehrmacht and Luftwaffe. Issued with a set of five rotors, from I to V, and two reflectors, B and C."
Case 2
    '3-rotor M3 Kriegsmarine
    Me.lblDescription.Caption = "This is the 3 rotor Kriegsmarine M3 Enigma machine, also called Funkschlussel M. Issued with a set of eight rotors, from I to VIII, and two reflectors, B and C."
Case 3
    '4-rotor M4 Kriegsmarine
    Me.lblDescription.Caption = "This is the 4 rotor Kriegsmarine M4 Enigma machine. Issued with a set of eight rotors, from I to VIII, two special thin rotors called Beta and Gamma (which don't advance), and two thin reflectors, B and C."
Case 4
    '4-rotor M4 Kriegsmarine (M3 and Wehrmacht/Luftwaffe Compatible)
    Me.lblDescription.Caption = "This is the 4 rotor Kriegsmarine M4 Enigma machine in compatible configuration for communication with M3 or Wehrmacht/Luftwaffe models.. Uses only the first five rotors, from I to V, reflector B together with Beta rotor or reflector C together with Gamma rotor. Beta and Gamma must have ringsetting A and remain in startposition A, to be compatible."
Case 5
    '4-rotor M4 Kriegsmarine (M3 Compatible)
    Me.lblDescription.Caption = "This is the 4 rotor Kriegsmarine M4 Enigma machine in compatible configuration for communication with M3 models. Uses all issued rotors, from I to VIII, reflector B together with Beta rotor or reflector C together with Gamma rotor. Beta and Gamma must have ringsetting A and remain in startposition A, to be compatible with M3 model."
End Select
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdSelect_Click()
'select default folder
Dim tmp As String
tmp = Browse("Select default Folder to Save...")
If tmp <> "" Then
    gstrSaveFolder = tmp
    Me.lblFolder.Caption = TrimPath(tmp, Me.lblFolder.Width)
    End If
End Sub

Private Sub cmdCreate_Click()
'create codebook
Dim fileO As Integer
Dim strFileName As String
Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim sp As Integer

Dim tmp As String
Dim intMonth As Integer
Dim M4model As Boolean
Dim Compatible As Boolean
Dim intRotorChoice As Integer
Dim intTotalSheets As Integer
Dim strPairs As String
Dim strCurrentMonth As String
Dim intCurrentMonth As Integer

'check codename
tmp = Trim(Me.txtCodeName.Text)
If tmp = "" Then
    MsgBox "Please enter Codename Net.", vbCritical
    Exit Sub
    End If
If InStr(1, tmp, "\") Or InStr(1, tmp, "/") Or InStr(1, tmp, ":") Or InStr(1, tmp, "*") Or InStr(1, tmp, "?") Or _
    InStr(1, tmp, Chr(34)) Or InStr(1, tmp, "<") Or InStr(1, tmp, ">") Or InStr(1, tmp, "|") Then
    MsgBox "Following characters are not allowed in Codename: \ / : * ? " & Chr(34) & " < > |", vbCritical
    Exit Sub
    End If

'check year
If Val(Me.txtYear.Text) < 1000 Or Val(Me.txtYear.Text) > 2999 Then
    MsgBox "Please enter a valid year.", vbCritical
    Exit Sub
    End If

Screen.MousePointer = 11
Me.frameCodebook.Enabled = False
Me.frameModel.Enabled = False
Me.ProgressBar.Value = 0
gstrCodeBook = ""

'select model
Select Case Me.cmbModel.ListIndex + 1
Case 1
    '3-rotor Wehrmacht/Luftwaffe
    M4model = False
    Compatible = False
    intRotorChoice = 5
Case 2
    '3-rotor M3 Kriegsmarine
    M4model = False
    Compatible = False
    intRotorChoice = 8
Case 3
    '4-rotor M4 Kriegsmarine
    M4model = True
    Compatible = False
    intRotorChoice = 8
Case 4
    '4-rotor M4 Kriegsmarine (M3 and Wehrmacht/Luftwaffe Compatible)
    M4model = True
    Compatible = True
    intRotorChoice = 5
Case 5
    '4-rotor M4 Kriegsmarine (M3 Compatible)
    M4model = True
    Compatible = True
    intRotorChoice = 8
End Select

'selec 1 sheet/whole year
intMonth = Me.cmbMonth.ListIndex
If intMonth <> 0 Then
    intTotalSheets = 1
    Else
    intTotalSheets = 12
    Me.ProgressBar.Visible = True
    End If

Randomize
'compose codebook sheet(s)
For i = 1 To intTotalSheets
    Me.ProgressBar.Value = i
    If intMonth <> 0 Then
        'one sheet
        strCurrentMonth = gstrMonthName(intMonth)
        intCurrentMonth = intMonth
            tmp = Trim(Str(intMonth)): If Len(tmp) = 1 Then tmp = "0" & tmp
        strFileName = Me.txtCodeName.Text & " " & Me.txtYear.Text & " " & tmp & ".txt"
        Else
        '12 sheets
        strCurrentMonth = gstrMonthName(i)
        intCurrentMonth = i
        tmp = Trim(Str(i)): If Len(tmp) = 1 Then tmp = "0" & tmp
        strFileName = Me.txtCodeName.Text & " " & Me.txtYear.Text & " " & tmp & ".txt"
        End If
    
    gstrCodeBook = vbCrLf
    If M4model Then gstrCodeBook = gstrCodeBook & "   "
    gstrCodeBook = gstrCodeBook & " GEHEIM!                SONDER-MASCHINENSCHLUSSEL: "
    tmp = Left(Trim(Me.txtCodeName.Text), 15)
    gstrCodeBook = gstrCodeBook & tmp
    If Not M4model Then
        gstrCodeBook = gstrCodeBook & Space(37 - Len(tmp) - Len(strCurrentMonth)) & strCurrentMonth & " " & Me.txtYear & vbCrLf & vbCrLf
        Else
        gstrCodeBook = gstrCodeBook & Space(29 - Len(tmp) - Len(strCurrentMonth)) & strCurrentMonth & " " & Me.txtYear & vbCrLf & vbCrLf
        End If
    
    'set codesheet header
    If Not M4model Then
        gstrCodeBook = gstrCodeBook & " --------------------------------------------------------------------------------------------" & vbCrLf
        gstrCodeBook = gstrCodeBook & " |Tag |UKW|     Walzenlage   |Ringstellung|      Steckerverbindungen      |   Kenngruppen   |" & vbCrLf
        gstrCodeBook = gstrCodeBook & " --------------------------------------------------------------------------------------------" & vbCrLf
        Else
        gstrCodeBook = gstrCodeBook & "    ------------------------------------------------------------------------------------" & vbCrLf
        gstrCodeBook = gstrCodeBook & "    |Tag |UKW|        Walzenlage       | Ringstellung  |      Steckerverbindungen      |" & vbCrLf
        gstrCodeBook = gstrCodeBook & "    ------------------------------------------------------------------------------------" & vbCrLf
        End If
        
    'loop thourgh all days of a month - reverse order!!!
    For k = gintDaysInMonth(intCurrentMonth) To 1 Step -1
        tmp = Trim(Str(k))
        If Len(tmp) = 1 Then tmp = "0" & tmp
        If M4model Then gstrCodeBook = gstrCodeBook & "   "
        gstrCodeBook = gstrCodeBook & " | " & tmp & " | "
        'get reflectors and 4th rotor
        gstrCodeBook = gstrCodeBook & GetRefAndFourth(M4model, Compatible)
        'get normal rotors
        gstrCodeBook = gstrCodeBook & "  " & GetRotors(intRotorChoice)
        'get rings
        gstrCodeBook = gstrCodeBook & " |  " & GetRings(M4model, Compatible) & " | "
        'get plugs
        strPairs = ""
        For j = 1 To 10
            tmp = GetNewPair(strPairs)
            Call InsertPair(tmp, strPairs)
        Next
        gstrCodeBook = gstrCodeBook & strPairs & "|"
       'get kenngruppen
        If Not M4model Then
            gstrCodeBook = gstrCodeBook & " " & GetKennGruppen & "|" & vbCrLf
            Else
            gstrCodeBook = gstrCodeBook & vbCrLf
            End If
    Next k
    If Not M4model Then
        gstrCodeBook = gstrCodeBook & " --------------------------------------------------------------------------------------------" & vbCrLf
        Else
        gstrCodeBook = gstrCodeBook & "    ------------------------------------------------------------------------------------" & vbCrLf
        End If
    'add remarks for compatible models
    If Me.cmbModel.ListIndex + 1 = 4 Then
        gstrCodeBook = gstrCodeBook & "    ACHTUNG! UKW B mit Beta  auf startposition A = M3 oder Wehrmacht/Luftwaffe mit UKW B" & vbCrLf
        gstrCodeBook = gstrCodeBook & "             UKW C mit Gamma auf startposition A = M3 oder Wehrmacht/Luftwaffe mit UKW C" & vbCrLf
    ElseIf Me.cmbModel.ListIndex + 1 = 5 Then
        gstrCodeBook = gstrCodeBook & "    ACHTUNG! UKW B mit Beta  auf startposition A = M3 mit UKW B" & vbCrLf
        gstrCodeBook = gstrCodeBook & "             UKW C mit Gamma auf startposition A = M3 mit UKW C" & vbCrLf
        End If
    
    'save codebook
    strFileName = gstrSaveFolder & strFileName
    On Error GoTo errHandler
    Screen.MousePointer = 11
    fileO = FreeFile
    Open strFileName For Output As #fileO
    Print #fileO, gstrCodeBook
    Close #fileO

Next i

Me.frameCodebook.Enabled = True
Me.frameModel.Enabled = True
Me.ProgressBar.Value = 0
Me.ProgressBar.Visible = False
Screen.MousePointer = 0
Me.Hide

If intMonth <> 0 Then
    frmMain.txtCodeBook.Text = gstrCodeBook
    Else
    frmMain.txtCodeBook.Text = vbCrLf & " Succesfully created 12 codesheets named " & Me.txtCodeName.Text & vbCrLf & vbCrLf & " The codesheets are saved in " & gstrSaveFolder
    End If
    
If gblnPrinterPresent = True And gstrCodeBook <> "" Then
    frmMain.mnuPrint.Enabled = True
    Else
    frmMain.mnuPrint.Enabled = True
    End If

Exit Sub
errHandler:
If Err Then MsgBox "Failed saving codesheet." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical
frmMain.txtCodeBook.Text = vbCrLf & " Failed creating codesheets!"
Screen.MousePointer = 0
Me.frameCodebook.Enabled = True
Me.frameModel.Enabled = True
Me.ProgressBar.Value = 0
Me.ProgressBar.Visible = False
End Sub

