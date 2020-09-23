VERSION 5.00
Begin VB.Form frmAlternative1 
   Caption         =   "Who needs passwords anyway?"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblReady 
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Click the labels above in the correct sequence  (innocent riddle) and then click the close button(top right)"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "a password!"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblAccess 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "access without"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblHowto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "How to"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmAlternative1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' In this example three objects must be clicked in a predefined sequence, followed
' by a click on the close button.
' For demonstration purposes I included the riddle.
' Variations on how many objects to click, the type of objects to click, how many
' times,... unlimited possibilities.
Dim PassValue As Integer

Private Sub cmdExit_Click()   ' normal exit
   Unload Me
End Sub

Private Sub cmdNext_Click()
   frmAlternative2.Show
   PassValue = 9     ' just a precaution
   Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If PassValue = 3 Then       ' correct sequence
      Cancel = 1               ' do not unload
      lblReady.Visible = True  ' tell user it's OK
      cmdNext.Visible = True   ' make next example available
      PassValue = 9            ' enable close button
      End If
End Sub

Private Sub lblAccess_Click()
   PassValue = IIf(PassValue = 1, 2, 9)
End Sub

Private Sub lblHowto_Click()
   PassValue = IIf(PassValue = 0, 1, 9)
End Sub

Private Sub lblPassword_Click()
   PassValue = IIf(PassValue = 2, 3, 9)
End Sub
