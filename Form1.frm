VERSION 5.00
Object = "{05589FA0-C356-11CE-BF01-00AA0055595A}#2.0#0"; "AMOVIE.OCX"
Begin VB.Form Form1 
   Caption         =   "Cd Test"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   495
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open From CD"
      Height          =   555
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin AMovieCtl.ActiveMovie mov 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      AutoStart       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo error
Dim x, y As Variant
x = Drive1.ListCount
y = x + 65
Text1.Text = Chr(y)
Drive1.Drive = Text1.Text
Dir1.Path = Drive1.Drive
Dir1.Path = "\movie"
mov.filename = Dir1.Path & "\cd iub4(SOUND).dat"
Exit Sub
error:
MsgBox "Please Insert The CD Rom..", 16, "No Cd Device"
End Sub

