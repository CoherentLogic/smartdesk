Attribute VB_Name = "modXOSWin"
Public JO As Boolean
Public InADrag As Boolean
Public ShowedTipOnce As Boolean
Public ActiveFrame As Integer
Public StartingAddress As String

Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Public WinText As String * 250
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

Public Const GW_HWNDNEXT = 2


   Sub LoadTaskList()
      'Get the hWnd of the first item in the master list
      'so we can process the task list entries (top-level only).
      
   End Sub

Public Sub Delay(Period As Integer)
    Start = Timer
    Do
    Loop Until Timer - Start >= Period
End Sub

Public Sub ReloadDesktop()
    frmBlank.Show
    frmMessage.Show
    frmMessage.Refresh
    
    FullQuit = False
    
    Unload frmToolbox
    Unload frmLockDesktop
    frmLockDesktop.Show
    
    Unload frmMessage
    Unload frmBlank
End Sub

Sub Main()
    If szXOSVar("MISC_SHOWINTRO") = "TRUE" Then
        frmXS9PIntro.Show vbModal
    End If
    ShowedTipOnce = False
    JO = False
    frmBlank.Show
    
    Delay Val(szXOSVar("MISC_LOGOTIMING"))
    If szXOSVar("POLICY_REQLOG") = "TRUE" Then
        frmAdminLog.Show
        frmAdminLog.SetFocus
        Do
            DoEvents
        Loop Until JO = True
        
    End If
    
    frmMessage.Label1(0).Caption = "Loading SmartDesk"
    frmMessage.Show
    frmMessage.Refresh
    frmLockDesktop.Show
    
    Unload frmMessage
    Unload frmBlank
   
    frmLockDesktop.WindowState = 2
End Sub


