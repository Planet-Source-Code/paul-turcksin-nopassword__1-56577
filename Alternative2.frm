VERSION 5.00
Begin VB.Form frmAlternative2 
   Caption         =   "Who needs passwords anyway?"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frPassword 
      Caption         =   "Enter your password   "
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   4215
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox fPassword 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Your secret password"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "and the end of this demonstration.   Thanks for trying it out.               Hope you enjoyed it."
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   2655
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
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Click the form, then the Minimize button, and the form again."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   $"Alternative2.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmAlternative2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' tricky : do NOT change or click anything in the frame: access is locked
' click form, click minimize and click form again...
Dim swNono As Boolean      ' clicked or changed in the password frame
Dim PassValue As Integer

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   swNono = True ' don't touch
End Sub

Private Sub Form_Click()
' if you couldn't resist touching ...
   If swNono Then Exit Sub
   
   If PassValue = 0 Then          ' first click
      PassValue = 1
   ElseIf PassValue = 2 Then      ' after minimize click
      frPassword.Visible = False  ' show ready and message
      End If
End Sub

Private Sub Form_Load()
' layout  (keep all form elements visible while in IDE, but not at run time)
   frPassword.Move 120, 1800
   Me.Height = 3555
End Sub

Private Sub Form_Resize()
' if this event occurs after a form click
   If PassValue = 1 And Me.WindowState = vbMinimized Then
      PassValue = 2
      Me.WindowState = vbNormal
'      Me.Show
      End If
'   DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAlternative2 = Nothing
End Sub

Private Sub fPassword_Change()
   swNono = True ' don't touch
End Sub

Private Sub fPassword_Click()
   swNono = True ' don't touch
End Sub

Private Sub fPassword_GotFocus()
' just to bait
   fPassword.SelStart = 0
   fPassword.SelLength = Len(fPassword)
End Sub
