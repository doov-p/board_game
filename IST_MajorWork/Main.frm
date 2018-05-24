VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000011&
   Caption         =   "Select a Game"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3840
   FillColor       =   &H0000FF00&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   800
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   2000
   End
   Begin VB.CommandButton btnGame2 
      Caption         =   "Hangman"
      Height          =   800
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2000
   End
   Begin VB.CommandButton btnGame1 
      Caption         =   "Tic Tac Toe"
      Height          =   800
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Setup audio function
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(strFileName As String)
sndPlaySound strFileName, 1
End Sub

'Check for exit button pressed
Private Sub btnExit_Click()
    PlaySound App.Path
    Unload Me
End Sub

'Tic Tac Toe pressed
Private Sub btnGame1_Click()
    frmGame1.Show
    Unload Me
End Sub

'Hangman pressed
Private Sub btnGame2_Click()
    frmGame2.Show
    Unload Me
End Sub

'Play music on form load
Private Sub Form_Load()
    PlaySound App.Path & "\Throne"
End Sub
