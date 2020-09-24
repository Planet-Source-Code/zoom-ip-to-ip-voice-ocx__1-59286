VERSION 5.00
Object = "*\AProject2.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1170
      TabIndex        =   2
      Top             =   3990
      Width           =   1245
   End
   Begin Project2.VOIP VOIP2 
      Height          =   2715
      Left            =   6330
      TabIndex        =   1
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   4789
   End
   Begin Project2.VOIP VOIP1 
      Height          =   2715
      Left            =   540
      TabIndex        =   0
      Top             =   780
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   4789
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VOIP1.connectit "127.0.0.1"
End Sub

Private Sub Form_Load()
VOIP1.setports 111, 222
VOIP2.setports 222, 111
End Sub
