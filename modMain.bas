Attribute VB_Name = "modMain"
'----------------------------------------------------------------
'
'    Enigma Codebook Tool
'
'    D. Rijmenants (c) 2005 dr.defcom@telenet.be
'
'----------------------------------------------------------------

Option Explicit

Public gstrSaveFolder       As String
Public gstrCurrentFolder    As String
Public gstrCodeBook         As String
Public gstrMonthName(12)    As String
Public gintDaysInMonth(12)  As Integer
Public ConvRot(8)           As String
Public dummy                As String

Public gblnPrinterPresent   As Boolean
Public gblnChangePageSetup  As Boolean
Public glngLeftMarginPrint  As Long
Public glngRightMarginPrint As Long
Public glngTopMarginPrint   As Long
Public glngBottMarginPrint  As Long

'Run or open file
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Browsing folders
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Sub Main()
Dim k As Integer

Load frmMain
Load frmCreate
Load frmAbout
Load frmPrintSetup

'get default folder
gstrSaveFolder = GetSetting(App.EXEName, "Config", "Folder", "c:\")
frmCreate.lblFolder.Caption = gstrSaveFolder
gstrCurrentFolder = gstrSaveFolder

'check if printer is present
On Error Resume Next
dummy = Printer.DeviceName
frmMain.mnuPrint.Enabled = False
If dummy <> "" And Err = 0 Then
    gblnPrinterPresent = True
    frmMain.mnuPageSetup.Enabled = True
    Else
    gblnPrinterPresent = False
    frmMain.mnuPageSetup.Enabled = False
    End If

'print page setup
glngTopMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintTop", "5"))
glngBottMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintBottom", "5"))
glngLeftMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintLeft", "5"))
glngRightMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintRight", "5"))

frmCreate.cmbModel.AddItem "3-rotor Wehrmacht/Luftwaffe"
frmCreate.cmbModel.AddItem "3-rotor M3 Kriegsmarine"
frmCreate.cmbModel.AddItem "4-rotor M4 Kriegsmarine"
frmCreate.cmbModel.AddItem "4-rotor M4 Kriegsmarine (M3 and Wehrmacht/Luftwaffe Compatible)"
frmCreate.cmbModel.AddItem "4-rotor M4 Kriegsmarine (M3 Compatible)"
frmCreate.cmbModel.ListIndex = 0

frmCreate.cmbMonth.AddItem "Complete year"
frmCreate.cmbMonth.AddItem "January"
frmCreate.cmbMonth.AddItem "February"
frmCreate.cmbMonth.AddItem "March"
frmCreate.cmbMonth.AddItem "April"
frmCreate.cmbMonth.AddItem "May"
frmCreate.cmbMonth.AddItem "June"
frmCreate.cmbMonth.AddItem "July"
frmCreate.cmbMonth.AddItem "August"
frmCreate.cmbMonth.AddItem "September"
frmCreate.cmbMonth.AddItem "October"
frmCreate.cmbMonth.AddItem "November"
frmCreate.cmbMonth.AddItem "December"
frmCreate.cmbMonth.ListIndex = 0

gstrMonthName(1) = "JANUAR"
gstrMonthName(2) = "FEBRUAR"
gstrMonthName(3) = "MARZ"
gstrMonthName(4) = "APRIL"
gstrMonthName(5) = "MAG"
gstrMonthName(6) = "JUNI"
gstrMonthName(7) = "JULI"
gstrMonthName(8) = "AUGUST"
gstrMonthName(9) = "SEPTEMBER"
gstrMonthName(10) = "OKTOBER"
gstrMonthName(11) = "NOVEMBER"
gstrMonthName(12) = "DEZEMBER"

gintDaysInMonth(1) = 31
gintDaysInMonth(2) = 29
gintDaysInMonth(3) = 31
gintDaysInMonth(4) = 30
gintDaysInMonth(5) = 31
gintDaysInMonth(6) = 30
gintDaysInMonth(7) = 31
gintDaysInMonth(8) = 31
gintDaysInMonth(9) = 30
gintDaysInMonth(10) = 31
gintDaysInMonth(11) = 30
gintDaysInMonth(12) = 31

ConvRot(1) = "I    "
ConvRot(2) = "II   "
ConvRot(3) = "III  "
ConvRot(4) = "IV   "
ConvRot(5) = "V    "
ConvRot(6) = "VI   "
ConvRot(7) = "VII  "
ConvRot(8) = "VIII "

frmMain.Show
End Sub

Public Function Browse(ByVal aTitle As String) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, pos%
Dim BrowsePath As String
Dim t
bInfo.hOwner = frmCreate.hwnd
bInfo.lpszTitle = aTitle
'the type of folder(s) to return
bInfo.ulFlags = &H1
'show the dialog box
pidl& = SHBrowseForFolder(bInfo)
'set the maximum characters
path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
'set the extracted path to SpecIn
Browse = Left(path$, pos - 1)
'make sure that "\" is at the end of the path
If Right$(Browse, 1) = "\" Then
    Browse = Browse
    Else
    Browse = Browse + "\"
End If
If Browse = "\" Then Browse = ""
End Function

Public Function TrimPath(ByVal Text As String, ByVal Size As Long)
'insert \...\ in file path if too large to fit in textbox
Dim TW
Dim Part1 As String
Dim Part2 As String
Dim pos As Integer
Size = Size - 1000
TW = frmCreate.picTextWidth.TextWidth(Text)
If TW < Size Then
    TrimPath = Text
    Exit Function
    End If
Part1 = Left(Text, 3) & "...\"
Part2 = Mid(Text, 4)
Text = Part1 & Part2
Do
TW = frmCreate.picTextWidth.TextWidth(Text)
If TW >= (Size) Then
    pos = InStr(1, Part2, "\")
    If pos <> 0 And pos < Len(Part2) Then
        Part2 = Mid(Part2, pos + 1)
        Else
        Part2 = Mid(Part2, 2)
        End If
    Text = Part1 & Part2
    End If
Loop While TW > (Size)
TrimPath = Text
End Function

Public Sub StartFile(ByVal FileName As String)
'execute a file
Dim RunCmd As String
Dim fExt As String
Dim x As Long
Dim RetS
On Error Resume Next
fExt = UCase(Right(FileName, 4))
Select Case fExt
Case ".WAV", ".MP2", ".MP3", ".MID", ".AVI"
    RunCmd = "play"
Case Else
    RunCmd = "open"
End Select
If FileExist(FileName) = False Or FileName = "" Then Exit Sub
'open file
RetS = ShellExecute(frmMain.hwnd, RunCmd, FileName, "", App.path, 1)
If RetS = 31 Then
    'if open fails, show the 'open with...'
    RetS = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & FileName)
    If RetS = 31 Then
        MsgBox "Unkwown filetype.", vbInformation
        End If
    End If
End Sub

Public Function FileExist(FileName As String) As Boolean
'checks weither a file exists
    On Error GoTo FileDoesNotExist
    Call FileLen(FileName)
    FileExist = True
    Exit Function
FileDoesNotExist:
    FileExist = False
End Function



