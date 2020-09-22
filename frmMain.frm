VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   " Enigma Codebook Tool"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11805
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodeBook 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuSheet 
      Caption         =   "&Codebook"
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select..."
      End
      Begin VB.Menu ln0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "&Create..."
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "&Page Setup..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
If Me.Width < 1000 Or Me.Height < 1000 Then Exit Sub
Me.txtCodeBook.Width = Me.Width - 100
Me.txtCodeBook.Height = Me.Height - 675
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (vbModal)
End Sub

Private Sub mnuHelp_Click()
Call StartFile(App.path & "\codebook.hlp")
End Sub

Private Sub mnuPageSetup_Click()
frmPrintSetup.Show (vbModal)
End Sub

Private Sub mnuSelect_Click()
'select a codebook
Dim strFileName As String
Dim fileO As Integer
Dim strInput As String

On Error Resume Next
With frmMain.comDlg
.FileName = ""
.DialogTitle = "Select Codebook..."
.Filter = "Text Files (*.txt)|*.txt"
.InitDir = gstrCurrentFolder
.FilterIndex = 1
.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
.ShowOpen
If Err = 32755 Then Exit Sub
strFileName = .FileName
gstrCurrentFolder = CurDir$
End With

'read file
On Error GoTo errHandler
Screen.MousePointer = 11
fileO = FreeFile
Open strFileName For Input As #fileO
    strInput = Input(LOF(fileO), 1)
Close #fileO
Screen.MousePointer = 0

If strInput <> "" Then
    gstrCodeBook = strInput
    frmMain.txtCodeBook.Text = gstrCodeBook
    End If
'update printer menu
If gblnPrinterPresent = True And gstrCodeBook <> "" Then
    Me.mnuPrint.Enabled = True
    Else
    Me.mnuPrint.Enabled = True
    End If

Exit Sub

errHandler:
MsgBox "Failed reading codebook." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical
End Sub

Private Sub mnuPrint_Click()
On Error Resume Next
Screen.MousePointer = 11
Printer.FontName = "Courier New"
Printer.FontSize = 12
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontStrikethru = False
Printer.Orientation = vbPRORLandscape 'vbPRORPortrait
Call PrintString(gstrCodeBook, glngLeftMarginPrint, glngRightMarginPrint, glngTopMarginPrint, glngBottMarginPrint)
Screen.MousePointer = 0
End Sub

Private Sub mnuCreate_Click()
frmCreate.Show (vbModal)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.EXEName, "Config", "Folder", gstrSaveFolder
Unload frmCreate
Unload frmAbout
Unload frmPrintSetup
End Sub


