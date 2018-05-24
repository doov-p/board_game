VERSION 5.00
Begin VB.Form frmGame1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn7 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   2000
   End
   Begin VB.CommandButton btn8 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   2000
   End
   Begin VB.CommandButton btn9 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   2000
   End
   Begin VB.CommandButton btn6 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton btn5 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton btn4 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton btn3 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   2000
   End
   Begin VB.CommandButton btn2 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2000
   End
   Begin VB.CommandButton btn1 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   2000
   End
   Begin VB.CommandButton btnMain 
      BackColor       =   &H0080C0FF&
      Caption         =   "Main Menu"
      Height          =   1000
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton btnReset 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      Height          =   1000
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1000
   End
   Begin VB.Label ywins 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label xwins 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label wins_y 
      BackColor       =   &H0080C0FF&
      Caption         =   "X wins:"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label wins_x 
      BackColor       =   &H0080C0FF&
      Caption         =   "O wins:"
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label turn 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label descturn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Current Turn:"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmGame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare public variables
Dim curPlr As Boolean
Dim turns As Integer
Dim win As Boolean
Dim player As String
Dim player2 As String
Dim win_x As Integer
Dim win_y As Integer

'Setup audio function
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(strFileName As String)
sndPlaySound strFileName, 1
End Sub

'Check if buttons are pressed
Private Sub btn1_Click()
    'Check if button blank
    If Me.btn1.Caption = "" Then
        If curPlr = True Then
            Me.btn1.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn1.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    'Disable button
    btn1.Enabled = False
End Sub

Private Sub btn2_Click()
    If Me.btn2.Caption = "" Then
        If curPlr = True Then
            Me.btn2.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn2.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn2.Enabled = False
End Sub

Private Sub btn3_Click()
    If Me.btn3.Caption = "" Then
        If curPlr = True Then
            Me.btn3.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn3.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn3.Enabled = False
End Sub

Private Sub btn4_Click()
    If Me.btn4.Caption = "" Then
        If curPlr = True Then
            Me.btn4.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn4.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn4.Enabled = False
End Sub

Private Sub btn5_Click()
    If Me.btn5.Caption = "" Then
        If curPlr = True Then
            Me.btn5.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn5.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn5.Enabled = False
End Sub

Private Sub btn6_Click()
    If Me.btn6.Caption = "" Then
        If curPlr = True Then
            Me.btn6.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn6.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn6.Enabled = False
End Sub

Private Sub btn7_Click()
    If Me.btn7.Caption = "" Then
        If curPlr = True Then
            Me.btn7.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn7.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn7.Enabled = False
End Sub

Private Sub btn8_Click()
    If Me.btn8.Caption = "" Then
        If curPlr = True Then
            Me.btn8.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn8.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn8.Enabled = False
End Sub

Private Sub btn9_Click()
    If Me.btn9.Caption = "" Then
        If curPlr = True Then
            Me.btn9.Caption = "X"
            curPlr = False
            turns = turns + 1
            turn.Caption = player2
            Call checkWin
        ElseIf curPlr = False Then
            Me.btn9.Caption = "O"
            curPlr = True
            turns = turns + 1
            turn.Caption = player
            Call checkWin
        End If
    End If
    btn9.Enabled = False
End Sub

'Check for winner
Public Sub checkWin()
    If curPlr = True Then
        If Me.btn1.Caption = "O" And Me.btn2.Caption = "O" And Me.btn3.Caption = "O" Or _
           Me.btn4.Caption = "O" And Me.btn5.Caption = "O" And Me.btn6.Caption = "O" Or _
           Me.btn7.Caption = "O" And Me.btn8.Caption = "O" And Me.btn9.Caption = "O" Or _
           Me.btn1.Caption = "O" And Me.btn4.Caption = "O" And Me.btn7.Caption = "O" Or _
           Me.btn2.Caption = "O" And Me.btn5.Caption = "O" And Me.btn8.Caption = "O" Or _
           Me.btn3.Caption = "O" And Me.btn6.Caption = "O" And Me.btn9.Caption = "O" Or _
           Me.btn1.Caption = "O" And Me.btn5.Caption = "O" And Me.btn9.Caption = "O" Or _
           Me.btn3.Caption = "O" And Me.btn5.Caption = "O" And Me.btn7.Caption = "O" Then
            'Disable all buttons
            Me.btn1.Enabled = False
            Me.btn2.Enabled = False
            Me.btn3.Enabled = False
            Me.btn4.Enabled = False
            Me.btn5.Enabled = False
            Me.btn6.Enabled = False
            Me.btn7.Enabled = False
            Me.btn8.Enabled = False
            Me.btn9.Enabled = False
            'Display win message
            MsgBox (player + " is the winner!")
            win = True
            win_x = win_x + 1
            xwins.Caption = win_x
            turns = 0
        End If
    ElseIf curPlr = False Then
        If Me.btn1.Caption = "X" And Me.btn2.Caption = "X" And Me.btn3.Caption = "X" Or _
           Me.btn4.Caption = "X" And Me.btn5.Caption = "X" And Me.btn6.Caption = "X" Or _
           Me.btn7.Caption = "X" And Me.btn8.Caption = "X" And Me.btn9.Caption = "X" Or _
           Me.btn1.Caption = "X" And Me.btn4.Caption = "X" And Me.btn7.Caption = "X" Or _
           Me.btn2.Caption = "X" And Me.btn5.Caption = "X" And Me.btn8.Caption = "X" Or _
           Me.btn3.Caption = "X" And Me.btn6.Caption = "X" And Me.btn9.Caption = "X" Or _
           Me.btn1.Caption = "X" And Me.btn5.Caption = "X" And Me.btn9.Caption = "X" Or _
           Me.btn3.Caption = "X" And Me.btn5.Caption = "X" And Me.btn7.Caption = "X" Then
            Me.btn1.Enabled = False
            Me.btn2.Enabled = False
            Me.btn3.Enabled = False
            Me.btn4.Enabled = False
            Me.btn5.Enabled = False
            Me.btn6.Enabled = False
            Me.btn7.Enabled = False
            Me.btn8.Enabled = False
            Me.btn9.Enabled = False
            MsgBox (player2 + " is the winner!")
            win = True
            win_y = win_y + 1
            ywins.Caption = win_y
            turns = 0
        End If
    End If
    'Check draw
    If turns = 9 And win = False Then
        Me.btn1.Enabled = False
        Me.btn2.Enabled = False
        Me.btn3.Enabled = False
        Me.btn4.Enabled = False
        Me.btn5.Enabled = False
        Me.btn6.Enabled = False
        Me.btn7.Enabled = False
        Me.btn8.Enabled = False
        Me.btn9.Enabled = False
        MsgBox ("Round Draw")
        turns = 0
    End If
End Sub

'Check if main menu pressed
Private Sub btnMain_Click()
    PlaySound App.Path
    frmMain.Show
    Unload Me
End Sub

'Check if reset pressed
Private Sub btnReset_Click()
    Me.btn1.Enabled = True
    Me.btn2.Enabled = True
    Me.btn3.Enabled = True
    Me.btn4.Enabled = True
    Me.btn5.Enabled = True
    Me.btn6.Enabled = True
    Me.btn7.Enabled = True
    Me.btn8.Enabled = True
    Me.btn9.Enabled = True
    
    Me.btn1.Caption = ""
    Me.btn2.Caption = ""
    Me.btn3.Caption = ""
    Me.btn4.Caption = ""
    Me.btn5.Caption = ""
    Me.btn6.Caption = ""
    Me.btn7.Caption = ""
    Me.btn8.Caption = ""
    Me.btn9.Caption = ""
    turns = 0
    win = False
    curPlr = False
End Sub

Private Sub Form_Load()
    PlaySound App.Path & "\Sakuya_Theme"
    'Get player names
    player = InputBox("Enter a name for player 1")
    If player = "" Then
        player = "Player 1"
    End If
    player2 = InputBox("Enter a name for player 2")
    If player2 = "" Then
        player2 = "Player 2"
    End If
    If player2 = player Then
        player2 = InputBox("Both players may not have the same name. Please reassign player 2.")
    End If
    
    'Reset functions and vars
    win_x = 0
    win_y = 0
    turns = 0
    Me.btn1.Enabled = True
    Me.btn2.Enabled = True
    Me.btn3.Enabled = True
    Me.btn4.Enabled = True
    Me.btn5.Enabled = True
    Me.btn6.Enabled = True
    Me.btn7.Enabled = True
    Me.btn8.Enabled = True
    Me.btn9.Enabled = True
    win = False
    curPlr = False
    turn.Caption = player2
End Sub
