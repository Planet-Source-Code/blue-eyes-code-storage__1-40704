VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "CodeUpdate"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      BackColor       =   &H00F2F4F4&
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
      Height          =   1800
      Left            =   285
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      ToolTipText     =   "Write your code here"
      Top             =   2310
      Width           =   6225
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00F2F4F4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6780
      MaxLength       =   50
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H80000016&
      Caption         =   "&Update"
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
      Height          =   390
      Left            =   5505
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4260
      Width           =   1050
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00F2F4F4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Function Name"
      Top             =   1200
      Width           =   1965
   End
   Begin VB.PictureBox Picture6 
      Height          =   5325
      Index           =   4
      Left            =   6750
      ScaleHeight     =   5265
      ScaleWidth      =   0
      TabIndex        =   11
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Picture6 
      Height          =   60
      Index           =   6
      Left            =   -495
      ScaleHeight     =   0
      ScaleWidth      =   7560
      TabIndex        =   10
      Top             =   4815
      Width           =   7620
   End
   Begin VB.PictureBox Picture6 
      Height          =   4845
      Index           =   2
      Left            =   -15
      ScaleHeight     =   4845
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6CCC2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6885
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   7
         Left            =   0
         ScaleHeight     =   1260
         ScaleWidth      =   15
         TabIndex        =   7
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture6 
         Height          =   1260
         Index           =   3
         Left            =   6750
         ScaleHeight     =   1260
         ScaleWidth      =   45
         TabIndex        =   6
         Top             =   0
         Width           =   45
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   7170
         TabIndex        =   5
         Top             =   -30
         Width           =   7230
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   1
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   6750
         TabIndex        =   4
         Top             =   435
         Width           =   6810
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H80000016&
         Cancel          =   -1  'True
         Caption         =   "X"
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
         Left            =   6390
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00D6CCC2&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Function"
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
         Left            =   105
         TabIndex        =   8
         Top             =   120
         Width           =   1725
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   300
      TabIndex        =   19
      Top             =   4470
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marked information is mandatory."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   4455
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write down your code below:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   285
      TabIndex        =   17
      Top             =   1980
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4875
      TabIndex        =   15
      Top             =   1290
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   1170
      TabIndex        =   13
      Top             =   1245
      Width           =   1590
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Here you can update the selected function."
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
      Height          =   270
      Left            =   300
      TabIndex        =   12
      Top             =   615
      Width           =   3825
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancle_Click()
    rtUpdate = False
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim strSQL As String
    If Len(Trim(txtCode)) <= 0 And Len(Trim(txtName)) <= 0 Then
        Load frmMsgBox
        frmError.lblErr.Caption = "You have to provide a function name and the code for that function. Please provide a function name and the code for the function."
        frmError.Visible = True
        Exit Sub
    ElseIf Len(Trim(txtCode)) <= 0 Then
        Load frmMsgBox
        frmError.lblErr.Caption = "There is no code in the code box. Please provide your code."
        frmError.Visible = True
        Exit Sub
    ElseIf Len(Trim(txtName)) <= 0 Then
        Load frmMsgBox
        frmError.lblErr.Caption = "You have to provide a function name. Please provide a function name."
        frmError.Visible = True
        Exit Sub
    End If
    strSQL = "update tblCodeStore set FunctionName='" & Trim(txtName) & "', Code='" & Trim(txtCode) & "', DateUpdate=#" & Now() & "# where CodeStoreID=" & CLng(txtID)
    DB.Execute strSQL
    
    selectedlvwID = CLng(txtID)
    strFunctionName = txtName
    strCode = txtCode
    Unload Me
End Sub

Private Sub Form_Load()
    MakeTop (Me.hwnd)
End Sub
