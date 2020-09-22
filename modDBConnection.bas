Attribute VB_Name = "mSubClass"
Option Explicit
    Public DB               As New ADODB.Connection
    Public rtCancle         As Boolean
    Public rtYes            As Boolean
    Public rtNo             As Boolean
    'Public rtUpdate        As Boolean
    Public strFunctionName  As String
    Public strCode          As String
    Public strProvider      As String
    Public selectedlvwID    As Long
    
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOZORDER = &H4
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_SHOWWINDOW = &H40
    
    Public Const TOP = -1
    Public Const NONTOP = -2
    

    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'


Public Function ChkPath(Path_to_Check As String) As String
    ChkPath = Path_to_Check
    If Right(ChkPath, 1) <> "\" Then ChkPath = ChkPath & "\"
End Function

Public Sub MakeTop(frmhWnd As Long)
    Dim rt As Long
    rt = SetWindowPos(frmhWnd, TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
