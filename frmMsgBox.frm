VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Code Storage"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000011&
      Height          =   2325
      Index           =   5
      Left            =   0
      ScaleHeight     =   2325
      ScaleWidth      =   15
      TabIndex        =   11
      Top             =   30
      Width           =   15
   End
   Begin VB.PictureBox Picture6 
      Height          =   2370
      Index           =   2
      Left            =   5865
      ScaleHeight     =   2310
      ScaleWidth      =   0
      TabIndex        =   10
      Top             =   15
      Width           =   60
   End
   Begin VB.CommandButton cmdNO 
      BackColor       =   &H80000016&
      Cancel          =   -1  'True
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1905
      Width           =   960
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H80000016&
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1905
      Width           =   960
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6CCC2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   -60
      ScaleHeight     =   555
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   -75
      Width           =   5955
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H80000016&
         Caption         =   "X"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5565
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   165
         Width           =   315
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   1
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   6105
         TabIndex        =   8
         Top             =   510
         Width           =   6165
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   6105
         TabIndex        =   7
         Top             =   45
         Width           =   6165
      End
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   3
         Left            =   5910
         ScaleHeight     =   1260
         ScaleWidth      =   30
         TabIndex        =   3
         Top             =   0
         Width           =   30
      End
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   7
         Left            =   60
         ScaleHeight     =   1260
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00D6CCC2&
         BackStyle       =   0  'Transparent
         Caption         =   "Code Storage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   60
      Index           =   6
      Left            =   -150
      ScaleHeight     =   0
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   2340
      Width           =   6165
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBox.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   720
      Left            =   300
      TabIndex        =   6
      Top             =   720
      Width           =   5100
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Private Sub cmdCancle_Click()
    rtYes = False
    rtNo = True
    rtCancle = True
    Unload Me
End Sub

Private Sub cmdNO_Click()
    rtYes = False
    rtNo = True
    rtCancle = True
    Unload Me
End Sub

Private Sub cmdYes_Click()
    rtYes = True
    rtNo = False
    rtCancle = False
    Unload Me
End Sub

Private Sub Form_Load()
    MakeTop (Me.hwnd)
End Sub
