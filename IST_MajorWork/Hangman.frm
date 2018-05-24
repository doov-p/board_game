VERSION 5.00
Begin VB.Form frmGame2 
   BackColor       =   &H00404040&
   Caption         =   "Hangman"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "m"
      Height          =   500
      Index           =   25
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "n"
      Height          =   500
      Index           =   24
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "b"
      Height          =   500
      Index           =   23
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "v"
      Height          =   500
      Index           =   22
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "c"
      Height          =   500
      Index           =   21
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "x"
      Height          =   500
      Index           =   20
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "z"
      Height          =   500
      Index           =   19
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "l"
      Height          =   500
      Index           =   18
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "k"
      Height          =   500
      Index           =   17
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "j"
      Height          =   500
      Index           =   16
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "h"
      Height          =   500
      Index           =   15
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "g"
      Height          =   500
      Index           =   14
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "f"
      Height          =   500
      Index           =   13
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "d"
      Height          =   500
      Index           =   12
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "s"
      Height          =   500
      Index           =   11
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "a"
      Height          =   500
      Index           =   10
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "p"
      Height          =   500
      Index           =   9
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "o"
      Height          =   500
      Index           =   8
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "i"
      Height          =   500
      Index           =   7
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "u"
      Height          =   500
      Index           =   6
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "y"
      Height          =   500
      Index           =   5
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "t"
      Height          =   500
      Index           =   4
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "r"
      Height          =   500
      Index           =   3
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "e"
      Height          =   500
      Index           =   2
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "w"
      Height          =   500
      Index           =   1
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton kboard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      Caption         =   "q"
      Height          =   500
      Index           =   0
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   720
      MaskColor       =   &H80000013&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton btnNewWord 
      BackColor       =   &H0080FF80&
      Caption         =   "New Word"
      Height          =   615
      Left            =   720
      MaskColor       =   &H80000013&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnHint 
      BackColor       =   &H00FFFF00&
      Caption         =   "Hint"
      Height          =   615
      Left            =   1200
      MaskColor       =   &H80000013&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   6
      Left            =   8880
      TabIndex        =   36
      Top             =   1800
      Width           =   375
   End
   Begin VB.Line arm 
      BorderColor     =   &H8000000B&
      Index           =   1
      X1              =   1680
      X2              =   1440
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line arm 
      BorderColor     =   &H8000000B&
      Index           =   0
      X1              =   1680
      X2              =   1920
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line leg 
      BorderColor     =   &H8000000B&
      Index           =   1
      X1              =   1680
      X2              =   1440
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Line leg 
      BorderColor     =   &H8000000B&
      Index           =   0
      X1              =   1680
      X2              =   1920
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Line body 
      BorderColor     =   &H8000000B&
      X1              =   1680
      X2              =   1680
      Y1              =   960
      Y2              =   1560
   End
   Begin VB.Shape head 
      BorderColor     =   &H8000000B&
      Height          =   495
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   480
      Width           =   495
   End
   Begin VB.Line bar 
      BorderColor     =   &H8000000B&
      Index           =   0
      X1              =   3000
      X2              =   2640
      Y1              =   600
      Y2              =   240
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   5
      Left            =   8160
      TabIndex        =   35
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   4
      Left            =   7440
      TabIndex        =   34
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   3
      Left            =   6720
      TabIndex        =   33
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   2
      Left            =   6000
      TabIndex        =   32
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   31
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label displetter 
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Index           =   0
      Left            =   4560
      TabIndex        =   30
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label hint 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Line rope 
      BorderColor     =   &H8000000B&
      X1              =   1680
      X2              =   1680
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line bar 
      BorderColor     =   &H8000000B&
      Index           =   1
      X1              =   3000
      X2              =   1680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      X1              =   3960
      X2              =   720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line post 
      BorderColor     =   &H8000000B&
      X1              =   3000
      X2              =   3000
      Y1              =   2520
      Y2              =   240
   End
End
Attribute VB_Name = "frmGame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare public variables
Dim failedGuesses As Integer
Dim wordlist(15) As String
Dim hintlist(15) As String
Dim rNumber, wordlength As Integer
Dim letter As String

'Setup audio function
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(strFileName As String)
sndPlaySound strFileName, 1
End Sub

'Check if hint pressed
Private Sub btnHint_Click()
    hint.Caption = (hintlist(rNumber))
End Sub

'Check if exit pressed
Private Sub btnExit_Click()
    PlaySound App.Path
    frmMain.Show
    Unload Me
End Sub

'Check if new word pressed
Private Sub btnNewWord_Click()
    'Reset failures
    failedGuesses = 0
    hint.Caption = ""
    
    'Enable keyboard
    For i = 0 To 25
        kboard(i).Enabled = True
    Next i

    'Hide hangman
    post.Visible = False
    bar(0).Visible = False
    bar(1).Visible = False
    rope.Visible = False
    head.Visible = False
    body.Visible = False
    leg(0).Visible = False
    leg(1).Visible = False
    arm(0).Visible = False
    arm(1).Visible = False
    
    'Hide letters
    For i = 0 To 6
        displetter(i).Visible = False
        displetter(i).Caption = ""
    Next i
    
    'Gen random number
    Randomize
    rNumber = Int((15 * Rnd) + 1)
    wordlength = Len(wordlist(rNumber))
    
    For i = 0 To wordlength
        displetter(i).Visible = True
    Next i
End Sub

'Check for keyboard press
Private Sub kboard_click(Index As Integer)
    letter = kboard(Index).Caption
    kboard(Index).Enabled = False
    Call checkLetter
End Sub

'Define wordlists and hintlists
Private Sub Form_Load()
    wordlist(0) = ""
    wordlist(1) = "kayak"
    wordlist(2) = "azure"
    wordlist(3) = "equip"
    wordlist(4) = "abyss"
    wordlist(5) = "crypt"
    wordlist(6) = "cobweb"
    wordlist(7) = "absurd"
    wordlist(8) = "injury"
    wordlist(9) = "exotic"
    wordlist(10) = "topaz"
    wordlist(11) = "kiosk"
    wordlist(12) = "quiz"
    wordlist(13) = "larynx"
    wordlist(14) = "moon"
    wordlist(15) = "galaxy"
    
    hintlist(0) = ""
    hintlist(1) = "A one person canoe."
    hintlist(2) = "The colour of the sky on a clear day."
    hintlist(3) = "To supply with necessary items for a purpose."
    hintlist(4) = "A deep or seemingly bottomless chasm."
    hintlist(5) = "An underground room or vault beneath a church."
    hintlist(6) = "An old and dusty spiderweb."
    hintlist(7) = "Wildly unreasonable, or illogical."
    hintlist(8) = "The instance of being damaged, or wounded."
    hintlist(9) = "Originating in a distant foreign area."
    hintlist(10) = "A typically yellow-orange precious stone."
    hintlist(11) = "A small store where newspapers, drinks, etc are sold."
    hintlist(12) = "A test of knowledge, or the act of questioning."
    hintlist(13) = "The technical name of the voice box."
    hintlist(14) = "Celestial bodies that orbit planets."
    hintlist(15) = "An area containing many solar systems."
    PlaySound App.Path & "\Sakuya_Theme"
    Call btnNewWord_Click
End Sub

'Check if letter is accurate
Public Sub checkLetter()
    Dim iscorrect As Boolean
    Dim p, start As Integer
    start = 1
    
    'Add letter to word
    For i = 1 To wordlength
        p = InStr(start, wordlist(rNumber), letter, 1)
        If p > 0 Then
            iscorrect = True
            displetter(p).Caption = letter
        End If
        start = p + 1
    Next i
    
    If iscorrect = False Then Call hangman
    
    'If hangman is complete display loss
    If arm(0).Visible = True Then
        For i = 0 To 25
            kboard(i).Enabled = False
        Next i
        MsgBox ("You Lose!")
        Call btnNewWord_Click
    End If
    
    'Checkwin
    If wordlength = 4 Then
        If Not displetter(1) = "" And Not displetter(2) = "" And Not displetter(3) = "" And Not displetter(4) = "" Then
            MsgBox ("You Win!")
            Call btnNewWord_Click
        End If
    ElseIf wordlength = 5 Then
        If Not displetter(1) = "" And Not displetter(2) = "" And Not displetter(3) = "" And Not displetter(4) = "" And Not displetter(5) = "" Then
            MsgBox ("You Win!")
            Call btnNewWord_Click
        End If
    ElseIf wordlength = 6 Then
        If Not displetter(1) = "" And Not displetter(2) = "" And Not displetter(3) = "" And Not displetter(4) = "" And Not displetter(5) = "" And Not displetter(6) = "" Then
            MsgBox ("You Win!")
            Call btnNewWord_Click
        End If
    End If
End Sub

'Draw hangman
Private Sub hangman()
    failedGuesses = failedGuesses + 1
    
    Select Case failedGuesses
        Case 1
            post.Visible = True
        Case 2
            bar(0).Visible = True
            bar(1).Visible = True
        Case 3
            rope.Visible = True
        Case 4
            head.Visible = True
        Case 5
            body.Visible = True
        Case 6
            leg(0).Visible = True
            leg(1).Visible = True
        Case 7
            arm(0).Visible = True
            arm(1).Visible = True
    End Select
End Sub
