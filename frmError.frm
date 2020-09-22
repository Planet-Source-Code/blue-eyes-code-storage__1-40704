VERSION 5.00
Begin VB.Form frmError 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Code Storage"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5865
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000011&
      Height          =   1890
      Index           =   4
      Left            =   0
      ScaleHeight     =   1890
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   -15
      Width           =   15
   End
   Begin VB.PictureBox Picture6 
      Height          =   1755
      Index           =   2
      Left            =   5850
      ScaleHeight     =   1695
      ScaleWidth      =   0
      TabIndex        =   8
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Picture6 
      Height          =   60
      Index           =   0
      Left            =   -90
      ScaleHeight     =   0
      ScaleWidth      =   5910
      TabIndex        =   6
      Top             =   1710
      Width           =   5970
      Begin VB.PictureBox Picture6 
         Height          =   90
         Index           =   1
         Left            =   0
         ScaleHeight     =   30
         ScaleWidth      =   5910
         TabIndex        =   7
         Top             =   0
         Width           =   5970
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   60
      Index           =   5
      Left            =   -60
      ScaleHeight     =   0
      ScaleWidth      =   5910
      TabIndex        =   5
      Top             =   375
      Width           =   5970
   End
   Begin VB.PictureBox Picture6 
      Height          =   60
      Index           =   6
      Left            =   -60
      ScaleHeight     =   0
      ScaleWidth      =   6105
      TabIndex        =   4
      Top             =   -30
      Width           =   6165
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
      TabIndex        =   2
      Top             =   -180
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
         Left            =   5550
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   7
         Left            =   60
         ScaleHeight     =   1260
         ScaleWidth      =   15
         TabIndex        =   11
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   3
         Left            =   5910
         ScaleHeight     =   1260
         ScaleWidth      =   30
         TabIndex        =   9
         Top             =   0
         Width           =   30
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
         TabIndex        =   3
         Top             =   270
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000016&
      Caption         =   "&OK"
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
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1305
      Width           =   960
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmError.frx":0000
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
      Left            =   345
      TabIndex        =   1
      Top             =   570
      Width           =   5100
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MakeTop (Me.hwnd)
End Sub
