VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who needs passwords anyway?"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Timer tmrPassword 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton Options 
      Caption         =   "Movies"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton Options 
      Caption         =   "Pictures"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Click options, checkbox, Browse : it all works!. But you cannot open the file.   By the way,. Exit works too!"
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   4095
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
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   $"Main.frx":0000
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select and open file:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' When this form is shown everything looks "normal" except for the fact that you
' cannot open the file selected.
' Please read comments in  cmdExit_Click() and tmrPassword_Timer()
Dim swAccess As Boolean            ' access control (user has to click Exit twice to get access to program)

Private Sub Check1_Click()
'  flag on/off will still shows, eventually confusing the user
End Sub

Private Sub cmdBrowse_Click()
   With CommonDialog1
      .ShowOpen
      lblFile = .FileName
      End With
End Sub

Private Sub cmdExit_Click()
' When application starts, tmrPass is disabled and swAccess is False.
' If timer is enabled then this event occured within .5 seconds and we can
' grant access.
   If tmrPassword.Enabled Then
      If swAccess = False Then    ' second click within .5 seconds
         tmrPassword.Enabled = False  ' do not need timer anymore
         swAccess = True          ' grant access
         lblReady.Visible = True  ' signal user has access
         cmdNext.Visible = True   ' next example
         End If
' timer is not enabled. The action now depends on switch 'swAccess".
'  if it is false the event enable timer and wait for second click event.
' If this switch is true this means the user had access and now wants to exit the application.
   Else
      If swAccess Then
         Unload Me
      Else
         tmrPassword.Enabled = True
         End If
      End If
End Sub

Private Sub cmdNext_Click()
   frmAlternative1.Show
   Unload Me
End Sub

Private Sub cmdOpen_Click()
' prevent execution if user doesn't have access
   If Not swAccess Then Exit Sub
   
   MsgBox "Open file " & lblFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMain = Nothing
End Sub

Private Sub tmrPassword_Timer()
' The timer provides a .5 seconds window for a second click event of command
' button "cmdExit" which will set switch "swAccess" to True.
' In case this happened the user was granted access and we do not need the
' timer anymore. If this switch is false - no access granted) terminate the
' application, ie the user clicked Exit once and thats what happens.
   If swAccess Then
      tmrPassword.Enabled = False
   Else
      Unload Me
      End If
End Sub

