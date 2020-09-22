VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About Enigma Codebook Tool"
   ClientHeight    =   3690
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   700
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1074.236
   ScaleMode       =   0  'User
   ScaleWidth      =   1123.685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2220
      Left            =   105
      ScaleHeight     =   2160
      ScaleWidth      =   900
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   960
      Begin VB.Image Image1 
         Height          =   600
         Left            =   150
         Picture         =   "frmAbout.frx":0000
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   105
      TabIndex        =   4
      Top             =   2310
      Width           =   5790
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   362
      Left            =   4515
      TabIndex        =   0
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   1380
      Left            =   1260
      TabIndex        =   6
      Top             =   960
      Width           =   4635
   End
   Begin VB.Label lblWarning 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   105
      TabIndex        =   1
      Top             =   2520
      Width           =   4350
   End
   Begin VB.Label lblTitle 
      Caption         =   "Enigma Codebook Tool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1260
      TabIndex        =   2
      Top             =   120
      Width           =   4725
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 2.0"
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   600
      Width           =   4605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.lblDescription.Caption = "Codebook creator for 3 rotor Wehrmacht/Luftwaffe, 3 rotor Kriegsmarine M3 and 4 rotor Kriegsmarine M4  Enigma cipher machine models." & vbCrLf & vbCrLf & "Programming by Dirk Rijmenants" & vbCrLf & vbCrLf & "© DEFCOM 1999 - 2005"
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

