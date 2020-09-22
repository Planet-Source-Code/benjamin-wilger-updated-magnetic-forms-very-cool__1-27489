VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Close All"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Me"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Form 3"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Form2"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2001 Benjamin Wilger"
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MOVE ME! SIZE ME!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3060
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

'Recommend way to end your application.
'DONT use End, VB WILL CRASH!
Private Sub Command4_Click()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub

Private Sub Form_Load()
    'With this little Command you enable Form-Docking
    DockingStart Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'And with this you'll deactivate it!
    'That's all you need to do!
    DockingTerminate Me
End Sub
