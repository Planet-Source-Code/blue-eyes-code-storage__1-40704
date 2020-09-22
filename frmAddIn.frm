VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Code Storage"
   ClientHeight    =   5565
   ClientLeft      =   2175
   ClientTop       =   1815
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   7035
      Left            =   -75
      ScaleHeight     =   7035
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   -75
      Width           =   1800
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   8
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   29
         Top             =   30
         Width           =   4050
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   4
         Left            =   -165
         ScaleHeight     =   0
         ScaleWidth      =   1905
         TabIndex        =   11
         Top             =   5595
         Width           =   1965
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   3
         Left            =   15
         ScaleHeight     =   0
         ScaleWidth      =   1905
         TabIndex        =   10
         Top             =   1245
         Width           =   1965
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   2
         Left            =   30
         ScaleHeight     =   0
         ScaleWidth      =   1905
         TabIndex        =   9
         Top             =   75
         Width           =   1965
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   2400
            Left            =   -1845
            TabIndex        =   13
            Top             =   -2385
            Width           =   3555
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   1
         Left            =   -90
         ScaleHeight     =   0
         ScaleWidth      =   1860
         TabIndex        =   8
         Top             =   5010
         Width           =   1920
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   0
         Left            =   -105
         ScaleHeight     =   0
         ScaleWidth      =   1860
         TabIndex        =   7
         Top             =   660
         Width           =   1920
         Begin VB.PictureBox Picture2 
            Height          =   3000
            Left            =   -1710
            ScaleHeight     =   2940
            ScaleWidth      =   3495
            TabIndex        =   12
            Top             =   -2970
            Width           =   3555
         End
      End
      Begin VB.PictureBox picAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6CCC2&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         ScaleHeight     =   525
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   5070
         Width           =   1815
         Begin VB.Label lblAbout 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00D6CCC2&
            BackStyle       =   0  'Transparent
            Caption         =   "About"
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
            Left            =   315
            TabIndex        =   6
            Top             =   135
            Width           =   645
         End
      End
      Begin VB.PictureBox picManagement 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6CCC2&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   -15
         ScaleHeight     =   525
         ScaleWidth      =   1815
         TabIndex        =   3
         Top             =   720
         Width           =   1815
         Begin VB.Label lblManagement 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00D6CCC2&
            BackStyle       =   0  'Transparent
            Caption         =   "Management"
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
            Left            =   285
            TabIndex        =   4
            Top             =   135
            Width           =   1350
         End
      End
      Begin VB.PictureBox picStore 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6CCC2&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         ScaleHeight     =   525
         ScaleWidth      =   1815
         TabIndex        =   1
         Top             =   135
         Width           =   1815
         Begin VB.Label lblStore 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00D6CCC2&
            BackStyle       =   0  'Transparent
            Caption         =   "Store"
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
            Left            =   285
            TabIndex        =   2
            Top             =   150
            Width           =   585
         End
      End
   End
   Begin VB.Frame fraAbout 
      Height          =   5610
      Left            =   1800
      TabIndex        =   26
      Top             =   -60
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Image Image1 
         Height          =   4440
         Left            =   3360
         Picture         =   "frmAddIn.frx":0000
         Stretch         =   -1  'True
         Top             =   915
         Width           =   4065
      End
      Begin VB.Label Label5 
         Caption         =   "Hello everybody, this is a add-in that can store useful code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   210
         TabIndex        =   34
         Top             =   345
         Width           =   7185
      End
   End
   Begin VB.Frame fraManagement 
      Height          =   5610
      Left            =   1800
      TabIndex        =   25
      Top             =   -60
      Visible         =   0   'False
      Width           =   7770
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   10
         Left            =   1605
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   41
         Top             =   1245
         Width           =   4050
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   9
         Left            =   1605
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   40
         Top             =   1545
         Width           =   4050
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   7
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   39
         Top             =   0
         Width           =   4050
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6CCC2&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1605
         ScaleHeight     =   240
         ScaleWidth      =   4035
         TabIndex        =   37
         Top             =   1305
         Width           =   4035
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00D6CCC2&
            BackStyle       =   0  'Transparent
            Caption         =   "Function(s)"
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
            Left            =   1200
            TabIndex        =   38
            Top             =   -15
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00D6CCC2&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3135
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4665
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00D6CCC2&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4140
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4665
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton cmdInsert 
         BackColor       =   &H00D6CCC2&
         Caption         =   "&Insert"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2145
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4665
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin MSComctlLib.ListView lvwMngFunction 
         Height          =   2895
         Left            =   1590
         TabIndex        =   28
         Top             =   1635
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00D6CCC2&
         BackStyle       =   0  'Transparent
         Caption         =   "Management"
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
         Left            =   105
         TabIndex        =   35
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lblManagementInfo 
         Caption         =   "The following codes are found. You can select any code and click Paste to paste the code in your form. Or just double click on it."
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
         Height          =   570
         Left            =   210
         TabIndex        =   27
         Top             =   600
         Width           =   7140
      End
      Begin VB.Label lblManInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "This code is submitted by Nazmul Alam Rubel on 4th November, 2002"
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
         Height          =   240
         Left            =   1740
         TabIndex        =   31
         Top             =   5280
         Width           =   5970
      End
   End
   Begin VB.Frame fraStore 
      Height          =   5610
      Left            =   1800
      TabIndex        =   15
      Top             =   -60
      Width           =   7770
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6CCC2&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1605
         ScaleHeight     =   240
         ScaleWidth      =   4035
         TabIndex        =   23
         Top             =   1305
         Width           =   4035
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00D6CCC2&
            BackStyle       =   0  'Transparent
            Caption         =   "Function(s)"
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
            Left            =   1200
            TabIndex        =   24
            Top             =   -15
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   6
         Left            =   1605
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   22
         Top             =   1245
         Width           =   4050
      End
      Begin VB.PictureBox Picture6 
         Height          =   60
         Index           =   5
         Left            =   1605
         ScaleHeight     =   0
         ScaleWidth      =   3990
         TabIndex        =   21
         Top             =   1545
         Width           =   4050
      End
      Begin MSComctlLib.ListView lvwCode 
         Height          =   840
         Left            =   8625
         TabIndex        =   20
         Top             =   3885
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1482
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00D6CCC2&
         Caption         =   "&Paste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3375
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4770
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin MSComctlLib.ListView lvwFunctionName 
         Height          =   2895
         Left            =   1590
         TabIndex        =   16
         Top             =   1635
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00D6CCC2&
         BackStyle       =   0  'Transparent
         Caption         =   "Stored Code"
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
         Left            =   105
         TabIndex        =   36
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "This code is submitted by Nazmul Alam Rubel on 4th November, 2002"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   18
         Top             =   5325
         Width           =   5970
      End
      Begin VB.Label lblTopInfo 
         Caption         =   "The following codes are found. You can select any code and click Paste to paste the code in your form. Or just double click on it."
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
         Height          =   570
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   7140
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   6900
      Left            =   1710
      TabIndex        =   14
      Top             =   15
      Width           =   60
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit
Dim SelectedLineNo As Long
Dim ItemClick As Boolean
Dim ManItemClick As Boolean
Dim isCodePage As Boolean
'



Private Sub cmdCopy_Click()
    Dim StartLine As Long, EndLine As Long
    Dim StartCol As Long, EndCol As Long
    Dim strData As String
    
    If ItemClick = False Then Exit Sub
    VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection StartLine, StartCol, EndLine, EndCol 'Getting the current line no
    strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1) ' getting the data of the current line
    
    Do While Len(strData) > 0
        EndLine = EndLine + 1
        strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    Loop
    
    ' Now you can past your selected code to the specific line
    SelectedLineNo = EndLine + 1 ' + 1 for another line space
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, "'    The following code is generated by Code Storage developed by Nazmul Alam Rubel"
    SelectedLineNo = SelectedLineNo + 1
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, "'    This code is submitted by " & lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(2) & " on " & Format(lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(3), "dddd, mmm d yyyy")
    SelectedLineNo = SelectedLineNo + 1
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(1)
    Unload Me
End Sub

Private Sub cmdDelete_Click()
   
    If ManItemClick = False Then Exit Sub
    Load frmMsgBox
    frmMsgBox.lblMessage = "Are you realy want to delete the function: " & lvwMngFunction.SelectedItem.Text
    frmMsgBox.Visible = True
    Do While frmMsgBox.Visible
        DoEvents
    Loop
    If rtYes Then
        DB.Execute "Delete from TblCodeStore where CodeStoreID=" & CLng(Trim(lvwCode.ListItems.Item(lvwMngFunction.SelectedItem.Index).Text))
        lvwCode.ListItems.Remove (lvwMngFunction.SelectedItem.Index)
        lvwMngFunction.ListItems.Remove (lvwMngFunction.SelectedItem.Index)
        rtYes = False
    Else
        rtNo = False
        rtCancle = False
    End If
    ManItemClick = False
End Sub

Private Sub cmdEdit_Click()
    If ManItemClick = False Then Exit Sub
    
    
    Load frmUpdate
    frmUpdate.Visible = True
    frmUpdate.txtName = lvwMngFunction.SelectedItem.Text
    frmUpdate.txtCode = lvwCode.ListItems.Item(lvwMngFunction.SelectedItem.Index).SubItems(1)
    frmUpdate.txtID = Trim(lvwCode.ListItems.Item(lvwMngFunction.SelectedItem.Index).Text)
    Me.Visible = False
    Do While frmUpdate.Visible = True
        DoEvents
    Loop
    Me.Visible = True
    Dim i As Long
    For i = 1 To lvwCode.ListItems.Count
        If lvwCode.ListItems.Item(i) = selectedlvwID Then
            lvwCode.ListItems.Item(i).SubItems(1) = strCode
            lvwMngFunction.ListItems.Item(i) = strFunctionName
            Exit For
        End If
    Next i
    
    
    ManItemClick = False
End Sub

Private Sub cmdInsert_Click()
    ManItemClick = False
    
    Me.Visible = False
    
    Load frmInsert
    frmInsert.Visible = True
    
    Do While frmInsert.Visible
        DoEvents
    Loop
    Me.Visible = True
    If rtCancle Then
        rtCancle = False
        Exit Sub
    End If
    lvwFunctionName.ListItems.Add , , strFunctionName
    lvwMngFunction.ListItems.Add , , strFunctionName
    Dim itmX As ListItem
    Set itmX = lvwCode.ListItems.Add(, , selectedlvwID)
    itmX.SubItems(1) = strCode
    itmX.SubItems(2) = strProvider
    itmX.SubItems(3) = Format(Now, "m/d/yyyy")
    
    
End Sub

Private Sub Form_Load()
    Dim StartLine As Long, EndLine As Long
    Dim StartCol As Long, EndCol As Long
    Dim strData As String
    Dim itmX As ListItem
    Dim i As Long
    
    MakeTop (Me.hwnd)
    lblInfo.Caption = ""
    ManItemClick = False
    
    lvwFunctionName.ColumnHeaders.Add , , "Function(s)", lvwFunctionName.Width
    lvwMngFunction.ColumnHeaders.Add , , "Function(s)", lvwFunctionName.Width
    lvwFunctionName.AllowColumnReorder = False
    ItemClick = False
    
    lvwCode.ColumnHeaders.Add , , "ID", lvwCode.Width / 4
    lvwCode.ColumnHeaders.Add , , "Code", lvwCode.Width / 4
    lvwCode.ColumnHeaders.Add , , "Provider", lvwCode.Width / 4
    lvwCode.ColumnHeaders.Add , , "Date", lvwCode.Width / 4
    
    Dim DT As New ADODB.Recordset
    DT.Open "Select * from TblCodeStore order by FunctionName", DB, adOpenStatic, adLockOptimistic
    
    If DT.RecordCount > 0 Then
        Do While DT.EOF = False
            Set itmX = lvwFunctionName.ListItems.Add(, , DT.Fields(1))
        
            Set itmX = lvwCode.ListItems.Add(, , DT.Fields(0))
            itmX.SubItems(1) = DT.Fields(2)
            If IsNull(DT.Fields(3)) Then
                itmX.SubItems(2) = ""
            Else
                itmX.SubItems(2) = DT.Fields(3)
            End If
            itmX.SubItems(3) = DT.Fields(4)
            DT.MoveNext
        Loop
        lblTopInfo.Caption = "The following codes are found. You can select any code and click Paste to paste the code in your form. Or just double click on it."
        lblManagementInfo.Caption = "The following codes are found. Now you can select any code to update or can insert your new function."
        fraStore.Visible = True
        fraManagement.Visible = False
        fraAbout.Visible = False
    Else
        lblTopInfo.Caption = "Sorry, no data found in the specific location. If you have already insert code and cannot get that code, please contact Nazmul Alam for more information."
        lblManagementInfo.Caption = "No code found in the Code Storage. However, here you can insert new function into Code Storage."
        fraStore.Visible = False
        fraManagement.Visible = True
        fraAbout.Visible = False
    End If
    Set DT = Nothing
    On Error GoTo ErrLevel
    VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection StartLine, StartCol, EndLine, EndCol 'Getting the current line no
    strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1) ' getting the data of the current line
    
    Do While Len(strData) > 0
        EndLine = EndLine + 1
        strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    Loop
    
    
    ' Now you can past your selected code to the specific line
    
    SelectedLineNo = EndLine + 1 ' + 1 for another line space
    
    
    'Set DB = Nothing
    isCodePage = True
    Exit Sub
ErrLevel:
    isCodePage = False
    Err.Clear
    
    Me.Visible = False
    Unload Me
    Load frmError
    frmError.Visible = True
    
End Sub

Private Sub Form_Terminate()
'Set DB = Nothing
End Sub

Private Sub lblAbout_Click()
    fraAbout.Visible = True
    fraStore.Visible = False
    fraManagement.Visible = False
End Sub

Private Sub lblManagement_Click()
    fraManagement.Visible = True
    fraStore.Visible = False
    fraAbout.Visible = False
    Dim itmX As ListItem
    Dim i As Long
    
    lblManInfo.Caption = ""
    
    ManItemClick = False
    lvwMngFunction.ListItems.Clear
    
    
    For i = 1 To lvwFunctionName.ListItems.Count
        Set itmX = lvwMngFunction.ListItems.Add(, , lvwFunctionName.ListItems.Item(i))
    Next i
    
End Sub

Private Sub lvwMngFunction_Click()
    ManItemClick = True
End Sub

Private Sub lblStore_Click()
    fraStore.Visible = True
    fraManagement.Visible = False
    fraAbout.Visible = False
    Dim StartLine As Long, EndLine As Long
    Dim StartCol As Long, EndCol As Long
    Dim strData As String
    Dim itmX As ListItem
    Dim i As Long
    
    lblInfo.Caption = ""
    
    

    lvwFunctionName.ListItems.Clear
    lvwCode.ListItems.Clear
    ItemClick = False
    

    Dim DT As New ADODB.Recordset
    
   
    DT.Open "Select * from TblCodeStore order by FunctionName", DB, adOpenStatic, adLockOptimistic
    
    If DT.RecordCount > 0 Then
        Do While DT.EOF = False
            Set itmX = lvwFunctionName.ListItems.Add(, , DT.Fields(1))
        
            Set itmX = lvwCode.ListItems.Add(, , DT.Fields(0))
            itmX.SubItems(1) = DT.Fields(2)
            If IsNull(DT.Fields(3)) Then
                itmX.SubItems(2) = ""
            Else
                itmX.SubItems(2) = DT.Fields(3)
            End If
            itmX.SubItems(3) = DT.Fields(4)
            DT.MoveNext
        Loop
        lblTopInfo.Caption = "The following codes are found. You can select any code and click Paste to paste the code in your form. Or just double click on it."
        lblManagementInfo.Caption = "The following codes are found. Now you can select any code to update or can insert your new function."
        fraStore.Visible = True
        fraManagement.Visible = False
        fraAbout.Visible = False
    Else
        lblTopInfo.Caption = "Sorry, no data found in the specific location. If you have already insert code and cannot get that code, please contact Nazmul Alam for more information."
        lblManagementInfo.Caption = "No code found in the Code Storage. However, here you can insert new function into Code Storage."
        fraStore.Visible = False
        fraManagement.Visible = True
        fraAbout.Visible = False
    End If
    On Error GoTo ErrLevel1
    VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection StartLine, StartCol, EndLine, EndCol 'Getting the current line no
    strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1) ' getting the data of the current line
    
    Do While Len(strData) > 0
        EndLine = EndLine + 1
        strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    Loop
    
    
    ' Now you can past your selected code to the specific line
    
    SelectedLineNo = EndLine + 1 ' + 1 for another line space
    
    
    Set DT = Nothing
    isCodePage = True
    Exit Sub
ErrLevel1:
    isCodePage = False
    Err.Clear
    Unload Me
    Load frmError
    frmError.Visible = True
    
End Sub

Private Sub lblStore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblStore
        .ForeColor = &HFF8080
        .FontUnderline = True
    End With
End Sub

Private Sub lvwFunctionName_Click()
    ItemClick = True
End Sub

Private Sub lvwFunctionName_DblClick()
    Dim StartLine As Long, EndLine As Long
    Dim StartCol As Long, EndCol As Long
    Dim strData As String
    
    VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection StartLine, StartCol, EndLine, EndCol 'Getting the current line no
    strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1) ' getting the data of the current line
    
    Do While Len(strData) > 0
        EndLine = EndLine + 1
        strData = VBInstance.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    Loop
    
    ' Now you can past your selected code to the specific line
    SelectedLineNo = EndLine + 1 ' + 1 for another line space
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, "'    The following code is generated by Code Storage developed by Nazmul Alam Rubel"
    SelectedLineNo = SelectedLineNo + 1
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, "'    This code is submitted by " & lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(2) & " on " & Format(lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(3), "dddd, mmm d yyyy")
    SelectedLineNo = SelectedLineNo + 1
    VBInstance.ActiveCodePane.CodeModule.InsertLines SelectedLineNo, lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(1)
    Unload Me
End Sub

Private Sub lvwFunctionName_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ItemClick = True
    lblInfo.Caption = "This code is submitted by " & lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(2) & " on " & Format(lvwCode.ListItems.Item(lvwFunctionName.SelectedItem.Index).SubItems(3), "dddd, mmm d yyyy")
End Sub

Private Sub picStore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblStore
        .ForeColor = &H800000
        .FontUnderline = False
    End With
End Sub
Private Sub lblManagement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblManagement
        .ForeColor = &HFF8080
        .FontUnderline = True
    End With
End Sub

Private Sub picManagement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblManagement
        .ForeColor = &H800000
        .FontUnderline = False
    End With
End Sub
Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblAbout
        .ForeColor = &HFF8080
        .FontUnderline = True
    End With
End Sub

Private Sub picAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblAbout
        .ForeColor = &H800000
        .FontUnderline = False
    End With
End Sub

