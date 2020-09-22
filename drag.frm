VERSION 5.00
Begin VB.Form drag 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblNumDropped 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "drag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
   List1.Clear
End Sub

Private Sub Command3_Click()
   End
End Sub

Private Sub Form_Load()

DragAcceptFiles Me.hWnd, True
End Sub
